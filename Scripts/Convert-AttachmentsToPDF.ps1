<#
.SYNOPSIS
    Finds every .pptx and .doc/.docx attachment in the vault that doesn't
    already have a matching PDF, converts it to PDF using PowerPoint/Word
    automation, and updates the note to embed the new PDF.

.DESCRIPTION
    This closes the gap where a recovered attachment only had an original
    file (pptx/doc) and no PDF twin - meaning it showed as a plain link
    instead of an embedded, scrollable preview. After this runs, every
    convertible attachment gets a PDF sitting next to it, and the note
    gets updated to embed it.

.NOTES
    Requires PowerPoint and/or Word to be installed and NOT past their
    deactivation date - flagged as time-sensitive given the "Product
    Deactivated" warning seen on this machine.

.EXAMPLE
    .\Convert-AttachmentsToPDF.ps1 -CourseName "Java Programming I (CIT 27000)"

.EXAMPLE
    .\Convert-AttachmentsToPDF.ps1
    (no -CourseName = scans every course under 01 - Courses)
#>

param(
    [string]$VaultCoursesPath = "C:\Users\dakota\Obsidian\Main Vault\School\01 - Courses",
    [string]$CourseName = ""
)

function Get-NormalizedName([string]$name) {
    $normalized = $name -replace '[\\/:\*\?"<>\|\-_]', ' '
    $normalized = $normalized -replace '\s+', ' '
    return $normalized.Trim()
}

$scanPath = if ($CourseName) { Join-Path $VaultCoursesPath $CourseName } else { $VaultCoursesPath }

if (-not (Test-Path $scanPath)) {
    Write-Host "Path not found: $scanPath" -ForegroundColor Red
    exit 1
}

# Find every pptx/doc/docx inside an Attachments folder that has no sibling PDF
$attachmentFiles = Get-ChildItem -Path $scanPath -Recurse -File -Include "*.pptx", "*.doc", "*.docx" |
    Where-Object { $_.Directory.Name -eq "Attachments" }

$toConvert = @()
foreach ($file in $attachmentFiles) {
    $pdfSibling = Join-Path $file.DirectoryName "$($file.BaseName).pdf"
    if (-not (Test-Path $pdfSibling)) {
        $toConvert += $file
    }
}

Write-Host "Found $($toConvert.Count) file(s) needing PDF conversion." -ForegroundColor Cyan
Write-Host ""

if ($toConvert.Count -eq 0) {
    Write-Host "Nothing to do." -ForegroundColor Green
    exit 0
}

$powerPoint = $null
$word = $null
$results = @()

foreach ($file in $toConvert) {
    $pdfPath = Join-Path $file.DirectoryName "$($file.BaseName).pdf"
    $isPptx = $file.Extension -eq ".pptx"

    # Run each conversion in an isolated job with a hard timeout, so a
    # single hung file (e.g. an invisible compatibility dialog) can only
    # ever cost a minute, never hours, like it did before this fix.
    $job = Start-Job -ScriptBlock {
        param($filePath, $pdfPath, $isPptx)
        try {
            if ($isPptx) {
                $app = New-Object -ComObject PowerPoint.Application
                $app.DisplayAlerts = 1
                $doc = $app.Presentations.Open($filePath, $true, $false, $false)
                $doc.SaveAs($pdfPath, 32)
                $doc.Close()
            } else {
                $app = New-Object -ComObject Word.Application
                $app.DisplayAlerts = 0
                $doc = $app.Documents.Open($filePath, $false, $true)
                $doc.SaveAs($pdfPath, 17)
                $doc.Close()
            }
            $app.Quit()
            return "OK"
        } catch {
            return "ERROR: $($_.Exception.Message)"
        }
    } -ArgumentList $file.FullName, $pdfPath, $isPptx

    $completed = Wait-Job -Job $job -Timeout 60

    if (-not $completed) {
        Write-Host "  [TIMEOUT] $($file.Name) took longer than 60 seconds - likely an invisible dialog. Killing it." -ForegroundColor Red
        Stop-Job -Job $job
        Remove-Job -Job $job -Force
        # Clean up any Office process the job may have left hung
        Get-Process -Name "WINWORD", "POWERPNT" -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue
        $results += [PSCustomObject]@{ File = $file.Name; Status = "Timed out after 60s"; Note = "" }
        continue
    }

    $jobResult = Receive-Job -Job $job
    Remove-Job -Job $job -Force

    if ($jobResult -ne "OK") {
        Write-Host "  [FAILED] $($file.Name): $jobResult" -ForegroundColor Red
        $results += [PSCustomObject]@{ File = $file.Name; Status = "Conversion failed"; Note = $jobResult }
        continue
    }

    Write-Host "  [converted] $($file.Name) -> $($file.BaseName).pdf" -ForegroundColor Green

    # Find and update the matching note to embed the new PDF
    $courseRoot = $file.FullName.Substring($VaultCoursesPath.Length).TrimStart('\').Split('\')[0]
    $courseRootPath = Join-Path $VaultCoursesPath $courseRoot
    $targetNormalized = Get-NormalizedName $file.BaseName
    $noteMatch = Get-ChildItem -Path $courseRootPath -Recurse -Filter "*.md" -File |
        Where-Object { (Get-NormalizedName $_.BaseName) -eq $targetNormalized } |
        Select-Object -First 1

    if ($noteMatch) {
        $noteContent = Get-Content -Path $noteMatch.FullName -Raw
        $embedLine = "![[$($file.BaseName).pdf]]"
        if ($noteContent -notmatch [regex]::Escape("$($file.BaseName).pdf")) {
            ($noteContent.TrimEnd() + "`r`n`r`n" + $embedLine + "`r`n") | Set-Content -Path $noteMatch.FullName -Encoding UTF8
            Write-Host "    embedded in: $($noteMatch.Name)" -ForegroundColor Green
        }
        $results += [PSCustomObject]@{ File = $file.Name; Status = "Converted and embedded"; Note = $noteMatch.Name }
    } else {
        $results += [PSCustomObject]@{ File = $file.Name; Status = "Converted, but no matching note found"; Note = "" }
    }
}

$reportPath = Join-Path $scanPath "_pdf-conversion-report.csv"
$results | Export-Csv -Path $reportPath -NoTypeInformation

Write-Host ""
$convertedCount = ($results | Where-Object { $_.Status -eq "Converted and embedded" }).Count
Write-Host "Done. Converted and embedded $convertedCount file(s)." -ForegroundColor Cyan
Write-Host "Report: $reportPath" -ForegroundColor Yellow
