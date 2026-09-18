<#
.SYNOPSIS
    Inserts a link into each Markdown note pointing to its recovered
    attachment file, using the original attachment-extraction report to
    know which note pairs with which file.

.DESCRIPTION
    The attachment recovery script copied real files (pptx/docx/pdf) into
    the vault, but never touched the notes themselves — so a note like
    "Values and Attitudes.md" still has no reference to the
    "KiniOB3e_C02_Final.pptx" sitting right next to it. This script reads
    the original report, finds each note by name, and inserts a link to
    its attachment(s) right after the title line.

    Safe to re-run — it checks for an existing link before adding one, so
    running this twice won't create duplicate links.

.NOTES
    Uses the ORIGINAL staging report (Purdue-Attachments\_attachment-
    extraction-report.csv), not the copy that landed inside the vault —
    that one may get deleted as part of cleanup.
#>

param(
    [string]$ReportPath           = "C:\Users\dakota\Documents\OneNoteExport\Purdue-Attachments\_attachment-extraction-report.csv",
    [string]$AttachmentsStagingPath = "C:\Users\dakota\Documents\OneNoteExport\Purdue-Attachments",
    [string]$VaultCoursesPath     = "C:\Users\dakota\Obsidian\Main Vault\School\01 - Courses"
)

if (-not (Test-Path $ReportPath)) {
    Write-Host "Report not found at: $ReportPath" -ForegroundColor Red
    exit 1
}

$rows = Import-Csv -Path $ReportPath | Where-Object { $_.Status -eq "Copied" }
Write-Host "Found $($rows.Count) recovered attachment(s) to link." -ForegroundColor Cyan
Write-Host ""

$results = @()

foreach ($row in $rows) {
    $stagingFile = $row.Detail
    $relative = $stagingFile.Substring($AttachmentsStagingPath.Length).TrimStart('\')
    $vaultAttachmentPath = Join-Path $VaultCoursesPath $relative

    if (-not (Test-Path $vaultAttachmentPath)) {
        Write-Host "  [skip] Attachment not found in vault: $relative" -ForegroundColor Yellow
        $results += [PSCustomObject]@{ Page = $row.Page; Status = "Attachment missing in vault"; Detail = $vaultAttachmentPath }
        continue
    }

    $noteFolder = Split-Path $vaultAttachmentPath -Parent
    $noteFileName = "$($row.Page).md"
    $notePath = Join-Path $noteFolder $noteFileName

    if (-not (Test-Path $notePath)) {
        # Fall back to a recursive search, but scoped to just this course's
        # own folder tree — searching the whole vault risks matching a
        # same-named page in a different course (e.g. two courses both
        # having a "Chapter 1 & 2 Homework" page)
        $found = Get-ChildItem -Path $noteFolder -Recurse -Filter $noteFileName -File -ErrorAction SilentlyContinue
        if ($found.Count -eq 1) {
            $notePath = $found[0].FullName
        } elseif ($found.Count -gt 1) {
            Write-Host "  [skip] Multiple notes named '$noteFileName' found within this course - ambiguous, handle manually" -ForegroundColor Yellow
            $results += [PSCustomObject]@{ Page = $row.Page; Status = "Ambiguous - multiple matches"; Detail = ($found.FullName -join "; ") }
            continue
        }
    }

    if (-not (Test-Path $notePath)) {
        Write-Host "  [skip] No matching note for: $($row.Page)" -ForegroundColor Yellow
        $results += [PSCustomObject]@{ Page = $row.Page; Status = "Note not found"; Detail = $notePath }
        continue
    }

    $attachmentFileName = Split-Path $vaultAttachmentPath -Leaf
    $linkLine = "Attachment: [[$attachmentFileName]]"

    $content = Get-Content -Path $notePath -Raw

    if ($content -match [regex]::Escape($attachmentFileName)) {
        Write-Host "  [already linked] $($row.Page) -> $attachmentFileName" -ForegroundColor DarkGray
        $results += [PSCustomObject]@{ Page = $row.Page; Status = "Already linked"; Detail = $attachmentFileName }
        continue
    }

    $lines = $content -split "`r?`n"
    if ($lines[0] -match "^#\s") {
        # Insert right after the title line
        $newLines = @($lines[0], "", $linkLine, "") + $lines[1..($lines.Length - 1)]
    } else {
        # No title line found — just prepend
        $newLines = @($linkLine, "") + $lines
    }

    ($newLines -join "`r`n") | Set-Content -Path $notePath -Encoding UTF8

    Write-Host "  [linked] $($row.Page) -> $attachmentFileName" -ForegroundColor Green
    $results += [PSCustomObject]@{ Page = $row.Page; Status = "Linked"; Detail = $attachmentFileName }
}

Write-Host ""
$linkedCount = ($results | Where-Object { $_.Status -eq "Linked" }).Count
Write-Host "Done. Linked $linkedCount note(s)." -ForegroundColor Cyan

$reportOut = Join-Path $AttachmentsStagingPath "_link-insertion-report.csv"
$results | Export-Csv -Path $reportOut -NoTypeInformation
Write-Host "Report: $reportOut" -ForegroundColor Yellow
