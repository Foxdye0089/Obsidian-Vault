<#
.SYNOPSIS
    Finds every page in a OneNote notebook with unrecoverable embedded
    images (callbackID=none), exports just that single page as its own
    small PDF directly into the matching vault folder, and links it into
    the note - replacing any garbled "[Image - no embedded data]" text
    with a clean link instead of leaving it there.

.DESCRIPTION
    This is the per-page version of the section-export approach: instead
    of one giant PDF per course that you have to search through, every
    affected page gets its own small PDF sitting right next to its note,
    linked directly from that note.

    Detection is done fresh against OneNote's live data (checking each
    page's raw XML for callbackID="none"), not from any previous report,
    so this catches everything - including pages missed by earlier scans.

.NOTES
    This can take a while with a large notebook - it's checking every
    page individually. Includes retry logic per page so one crash
    doesn't cascade and fail everything after it, unlike the earlier
    section-export script.
#>

param(
    [Parameter(Mandatory = $true)]
    [string]$NotebookName,

    [Parameter(Mandatory = $true)]
    [string]$VaultCoursesPath
)

function New-OneNoteConnection {
    return New-Object -ComObject OneNote.Application
}

function Get-SafeFileName([string]$name) {
    return ($name -replace '[\\/:\*\?"<>\|]', '-')
}

$OneNote = New-OneNoteConnection

# Find the notebook
[string]$notebooksXml = ""
$OneNote.GetHierarchy("", 2, [ref]$notebooksXml)  # hsNotebooks = 2
[xml]$notebooksDoc = $notebooksXml
$ns = $notebooksDoc.DocumentElement.NamespaceURI
$nsMgr = New-Object System.Xml.XmlNamespaceManager($notebooksDoc.NameTable)
$nsMgr.AddNamespace("one", $ns)

$notebookNode = $notebooksDoc.SelectSingleNode("//one:Notebook[@name='$NotebookName']", $nsMgr)
if (-not $notebookNode) {
    Write-Host "Notebook '$NotebookName' not found." -ForegroundColor Red
    exit 1
}
$notebookId = $notebookNode.ID

# Get full hierarchy
[string]$fullXml = ""
$OneNote.GetHierarchy($notebookId, 4, [ref]$fullXml)  # hsPages = 4
[xml]$fullDoc = $fullXml
$nsMgr2 = New-Object System.Xml.XmlNamespaceManager($fullDoc.NameTable)
$nsMgr2.AddNamespace("one", $ns)

$pageNodes = $fullDoc.SelectNodes("//one:Page", $nsMgr2)
Write-Host "Checking $($pageNodes.Count) pages for unrecoverable images..." -ForegroundColor Cyan
Write-Host ""

$results = @()
$checkedCount = 0

foreach ($pageNode in $pageNodes) {
    $checkedCount++
    $pageId = $pageNode.ID
    $pageName = $pageNode.name

    # Build relative folder path from ancestor sections/section groups
    $pathParts = @()
    $ancestor = $pageNode.ParentNode
    while ($ancestor -and $ancestor.LocalName -ne "Notebook") {
        if ($ancestor.name) { $pathParts = @($ancestor.name) + $pathParts }
        $ancestor = $ancestor.ParentNode
    }
    $relativeFolder = ($pathParts -join "\")
    $courseFolder = if ($pathParts.Count -gt 0) { $pathParts[0] } else { "" }

    [string]$pageXml = ""
    try {
        $OneNote.GetPageContent($pageId, [ref]$pageXml, 0)
    } catch {
        $results += [PSCustomObject]@{ Page = $pageName; Status = "ERROR reading page"; Detail = $_.Exception.Message }
        continue
    }

    if ($pageXml -notmatch 'callbackID="none"') {
        continue  # page is fine, nothing to fix
    }

    Write-Host "  [found broken] $pageName" -ForegroundColor Yellow

    $destFolder = Join-Path $VaultCoursesPath $relativeFolder
    if (-not (Test-Path $destFolder)) {
        New-Item -ItemType Directory -Path $destFolder -Force | Out-Null
    }

    $safeName = Get-SafeFileName $pageName
    $pdfPath = Join-Path $destFolder "$safeName.pdf"

    $exported = $false
    for ($attempt = 1; $attempt -le 2; $attempt++) {
        try {
            $OneNote.Publish($pageId, $pdfPath, 3, "")
            $exported = $true
            break
        } catch {
            Write-Host "    attempt $attempt failed: $($_.Exception.Message)" -ForegroundColor Red
            Start-Sleep -Seconds 2
            $OneNote = New-OneNoteConnection
        }
    }

    if (-not $exported) {
        $results += [PSCustomObject]@{ Page = $pageName; Status = "PDF export failed after retry"; Detail = $pdfPath }
        continue
    }

    Write-Host "    exported -> $safeName.pdf" -ForegroundColor Green

    # Find the matching note - same folder first, then fall back to a
    # recursive search scoped to this course only (avoids cross-course
    # name collisions like two courses both having "Review Questions")
    $noteFileName = "$pageName.md"
    $notePath = Join-Path $destFolder $noteFileName

    if (-not (Test-Path $notePath) -and $courseFolder) {
        $courseRoot = Join-Path $VaultCoursesPath $courseFolder
        $found = Get-ChildItem -Path $courseRoot -Recurse -Filter $noteFileName -File -ErrorAction SilentlyContinue
        if ($found.Count -eq 1) {
            $notePath = $found[0].FullName
        } elseif ($found.Count -gt 1) {
            $results += [PSCustomObject]@{ Page = $pageName; Status = "PDF exported, but note is ambiguous"; Detail = $pdfPath }
            continue
        }
    }

    if (-not (Test-Path $notePath)) {
        $results += [PSCustomObject]@{ Page = $pageName; Status = "PDF exported, but no matching note found"; Detail = $pdfPath }
        continue
    }

    $linkLine = "Full page export: [[$safeName.pdf]]"
    $content = Get-Content -Path $notePath -Raw

    if ($content -match [regex]::Escape("$safeName.pdf")) {
        $results += [PSCustomObject]@{ Page = $pageName; Status = "PDF exported, already linked"; Detail = $notePath }
        continue
    }

    # Strip out any garbled "[Image - no embedded data...]" blocks left by
    # the earlier Markdown export, replacing them with the clean link
    $cleaned = [regex]::Replace($content, '\[Image - no embedded data.*?\]\r?\n?', '')
    $cleaned = [regex]::Replace($cleaned, '(\r?\n){3,}', "`r`n`r`n")  # collapse extra blank lines

    $lines = $cleaned -split "`r?`n"
    if ($lines[0] -match "^#\s") {
        $newLines = @($lines[0], "", $linkLine, "") + $lines[1..($lines.Length - 1)]
    } else {
        $newLines = @($linkLine, "") + $lines
    }

    ($newLines -join "`r`n") | Set-Content -Path $notePath -Encoding UTF8

    Write-Host "    linked into: $(Split-Path $notePath -Leaf)" -ForegroundColor Green
    $results += [PSCustomObject]@{ Page = $pageName; Status = "Exported and linked"; Detail = $notePath }
}

$reportPath = Join-Path $VaultCoursesPath "_broken-page-fix-report.csv"
$results | Export-Csv -Path $reportPath -NoTypeInformation

Write-Host ""
Write-Host "Checked $checkedCount pages." -ForegroundColor Cyan
$linkedCount = ($results | Where-Object { $_.Status -eq "Exported and linked" }).Count
$failedCount = ($results | Where-Object { $_.Status -like "*failed*" -or $_.Status -like "*ERROR*" }).Count
Write-Host "Fixed and linked: $linkedCount" -ForegroundColor Green
Write-Host "Needs manual attention: $failedCount" -ForegroundColor Yellow
Write-Host "Full report: $reportPath" -ForegroundColor Cyan
