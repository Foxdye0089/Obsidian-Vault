<#
.SYNOPSIS
    Bulk-extracts original file attachments from a OneNote notebook by
    reading each page's raw XML directly via COM and copying the locally
    cached copy of each inserted file.

.DESCRIPTION
    When you insert a file into OneNote, OneNote keeps a local cached
    copy of that file on disk and references it from the page's XML via
    the InsertedFile element's `pathCache` attribute. This script walks
    every page in a notebook, finds every InsertedFile reference, and
    copies the cached file to an output folder that mirrors the
    notebook's section/page structure.

    This does NOT depend on the notebook being writable, and does NOT
    rely on rendering an image of the attachment (which is what fails
    for "attachment printout" pages in most Markdown exporters). It
    grabs the real, original file — searchable PDF/PPTX, not a picture.

.NOTES
    Requires the OneNote desktop app (Microsoft 365) to be running.
    Test on a small notebook or single section first before running
    against something large like Purdue.

.EXAMPLE
    .\Extract-OneNoteAttachments.ps1 -ListNotebooks

    Lists exact notebook names as OneNote currently sees them (use this
    first if you're not sure whether it shows "Purdue" or "IUPUI").

.EXAMPLE
    .\Extract-OneNoteAttachments.ps1 -NotebookName "IUPUI" -OutputPath "C:\Users\dakota\Documents\OneNoteExport\Purdue-Attachments"
#>

param(
    [string]$NotebookName,
    [string]$OutputPath,
    [switch]$ListNotebooks
)

$OneNote = New-Object -ComObject OneNote.Application

# Get all notebooks (hsNotebooks = 2)
[string]$notebooksXml = ""
$OneNote.GetHierarchy("", 2, [ref]$notebooksXml)

[xml]$notebooksDoc = $notebooksXml
$ns = $notebooksDoc.DocumentElement.NamespaceURI
$nsMgr = New-Object System.Xml.XmlNamespaceManager($notebooksDoc.NameTable)
$nsMgr.AddNamespace("one", $ns)

if ($ListNotebooks) {
    Write-Host "Notebooks currently visible to OneNote:" -ForegroundColor Cyan
    $notebooksDoc.SelectNodes("//one:Notebook", $nsMgr) | ForEach-Object {
        Write-Host "  - $($_.name)"
    }
    exit 0
}

if (-not $NotebookName -or -not $OutputPath) {
    Write-Host "Usage: .\Extract-OneNoteAttachments.ps1 -NotebookName <name> -OutputPath <path>" -ForegroundColor Yellow
    Write-Host "   or: .\Extract-OneNoteAttachments.ps1 -ListNotebooks" -ForegroundColor Yellow
    exit 1
}

$notebookNode = $notebooksDoc.SelectSingleNode("//one:Notebook[@name='$NotebookName']", $nsMgr)

if (-not $notebookNode) {
    Write-Host "Notebook '$NotebookName' not found. Run with -ListNotebooks to see exact names." -ForegroundColor Red
    exit 1
}

$notebookId = $notebookNode.ID

# Get full hierarchy (sections + pages) for this notebook (hsPages = 4)
[string]$fullXml = ""
$OneNote.GetHierarchy($notebookId, 4, [ref]$fullXml)

[xml]$fullDoc = $fullXml
$nsMgr2 = New-Object System.Xml.XmlNamespaceManager($fullDoc.NameTable)
$nsMgr2.AddNamespace("one", $ns)

$pageNodes = $fullDoc.SelectNodes("//one:Page", $nsMgr2)
Write-Host "Found $($pageNodes.Count) pages in '$NotebookName'." -ForegroundColor Cyan
Write-Host ""

$results = @()

foreach ($pageNode in $pageNodes) {
    $pageId = $pageNode.ID
    $pageName = $pageNode.name

    # Build a relative folder path from this page's ancestor sections/section groups
    $pathParts = @()
    $ancestor = $pageNode.ParentNode
    while ($ancestor -and $ancestor.LocalName -ne "Notebook") {
        if ($ancestor.name) { $pathParts = @($ancestor.name) + $pathParts }
        $ancestor = $ancestor.ParentNode
    }
    $relativeFolder = ($pathParts -join "\")

    [string]$pageXml = ""
    try {
        $OneNote.GetPageContent($pageId, [ref]$pageXml, 0)  # 0 = piBasic
    } catch {
        $results += [PSCustomObject]@{ Page = $pageName; Status = "ERROR reading page"; Detail = $_.Exception.Message }
        continue
    }

    [xml]$pageDoc = $pageXml
    $nsMgr3 = New-Object System.Xml.XmlNamespaceManager($pageDoc.NameTable)
    $nsMgr3.AddNamespace("one", $ns)

    $insertedFiles = $pageDoc.SelectNodes("//one:InsertedFile", $nsMgr3)
    if ($insertedFiles.Count -eq 0) { continue }

    $destFolder = Join-Path $OutputPath $relativeFolder
    New-Item -ItemType Directory -Path $destFolder -Force | Out-Null

    foreach ($file in $insertedFiles) {
        $pathCache = $file.pathCache
        $preferredName = $file.preferredName

        if ([string]::IsNullOrWhiteSpace($pathCache)) {
            $results += [PSCustomObject]@{ Page = $pageName; Status = "No pathCache attribute"; Detail = $preferredName }
            continue
        }

        if (-not (Test-Path $pathCache)) {
            $results += [PSCustomObject]@{ Page = $pageName; Status = "Cached file missing on disk"; Detail = $pathCache }
            continue
        }

        $destFile = Join-Path $destFolder $preferredName
        try {
            Copy-Item -Path $pathCache -Destination $destFile -Force
            $results += [PSCustomObject]@{ Page = $pageName; Status = "Copied"; Detail = $destFile }
            Write-Host "  [copied] $pageName -> $preferredName" -ForegroundColor Green
        } catch {
            $results += [PSCustomObject]@{ Page = $pageName; Status = "Copy failed"; Detail = $_.Exception.Message }
        }
    }
}

$reportPath = Join-Path $OutputPath "_attachment-extraction-report.csv"
$results | Export-Csv -Path $reportPath -NoTypeInformation

$copiedCount = ($results | Where-Object { $_.Status -eq "Copied" }).Count
$missingCount = ($results | Where-Object { $_.Status -eq "Cached file missing on disk" }).Count

Write-Host ""
Write-Host "Done. Copied $copiedCount attachment(s). $missingCount had no recoverable cached copy." -ForegroundColor Cyan
Write-Host "Full report: $reportPath" -ForegroundColor Yellow
