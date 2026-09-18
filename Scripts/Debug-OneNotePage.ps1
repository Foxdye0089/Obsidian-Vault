<#
.SYNOPSIS
    Dumps the raw XML OneNote returns for a single page, so we can see
    exactly how its content (including any attachments) is structured.

.EXAMPLE
    .\Debug-OneNotePage.ps1 -NotebookName "IUPUI" -CourseName "Security Risk Assessment" -PageName "Chapter 1 & 2 Homework"
#>

param(
    [Parameter(Mandatory = $true)]
    [string]$NotebookName,

    [Parameter(Mandatory = $true)]
    [string]$CourseName,

    [Parameter(Mandatory = $true)]
    [string]$PageName,

    [string]$OutputPath = "C:\Users\dakota\Downloads\page-debug.xml"
)

$OneNote = New-Object -ComObject OneNote.Application

[string]$notebooksXml = ""
$OneNote.GetHierarchy("", 2, [ref]$notebooksXml)
[xml]$notebooksDoc = $notebooksXml
$ns = $notebooksDoc.DocumentElement.NamespaceURI
$nsMgr = New-Object System.Xml.XmlNamespaceManager($notebooksDoc.NameTable)
$nsMgr.AddNamespace("one", $ns)

$notebookNode = $notebooksDoc.SelectSingleNode("//one:Notebook[@name='$NotebookName']", $nsMgr)
if (-not $notebookNode) {
    Write-Host "Notebook not found." -ForegroundColor Red
    exit 1
}
$notebookId = $notebookNode.ID

[string]$fullXml = ""
$OneNote.GetHierarchy($notebookId, 4, [ref]$fullXml)
[xml]$fullDoc = $fullXml
$nsMgr2 = New-Object System.Xml.XmlNamespaceManager($fullDoc.NameTable)
$nsMgr2.AddNamespace("one", $ns)

$allPageNodes = $fullDoc.SelectNodes("//one:Page", $nsMgr2)

$targetPage = $null
foreach ($pageNode in $allPageNodes) {
    if ($pageNode.name -ne $PageName) { continue }
    $ancestor = $pageNode.ParentNode
    $topAncestorName = $null
    while ($ancestor -and $ancestor.LocalName -ne "Notebook") {
        if ($ancestor.name) { $topAncestorName = $ancestor.name }
        $ancestor = $ancestor.ParentNode
    }
    if ($topAncestorName -eq $CourseName) {
        $targetPage = $pageNode
        break
    }
}

if (-not $targetPage) {
    Write-Host "Page '$PageName' not found under course '$CourseName'." -ForegroundColor Red
    exit 1
}

Write-Host "Found page. ID: $($targetPage.ID)" -ForegroundColor Cyan

[string]$pageXml = ""
$OneNote.GetPageContent($targetPage.ID, [ref]$pageXml, 0)

$pageXml | Out-File -FilePath $OutputPath -Encoding UTF8

Write-Host "Raw XML written to: $OutputPath" -ForegroundColor Green
Write-Host ""
Write-Host "Quick check - does it contain 'InsertedFile'?" -ForegroundColor Cyan
if ($pageXml -match "InsertedFile") {
    Write-Host "  YES - InsertedFile element(s) present" -ForegroundColor Green
    $matches = [regex]::Matches($pageXml, '<one:InsertedFile[^>]*>')
    foreach ($m in $matches) {
        Write-Host "  $($m.Value)" -ForegroundColor Yellow
    }
} else {
    Write-Host "  NO - no InsertedFile element found in this page's XML at all" -ForegroundColor Red
}
