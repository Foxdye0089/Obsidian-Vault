<#
.SYNOPSIS
    Exports one or more OneNote sections to PDF, using the same underlying
    mechanism as File > Export > Section > PDF in the OneNote UI.

.DESCRIPTION
    For pages where embedded images failed to export via the Markdown
    converter (callbackID=none — OneNote never generated a fetchable
    binary for them), this bypasses that broken path entirely by exporting
    full sections as PDF, which uses OneNote's normal rendering pipeline
    instead of the COM binary-fetch pipeline that's failing.

.NOTES
    Test with a single section first (the default $SectionNames below
    includes all 10 — trim it to one entry for your first run) before
    running the full batch.
#>

param(
    [Parameter(Mandatory = $true)]
    [string]$NotebookName,

    [Parameter(Mandatory = $true)]
    [string]$OutputPath,

    [string[]]$SectionNames = @(
        "Adv Network Administration (CIT 41500)",
        "Advanced Network Security (CIT 40600)",
        "Applied Secure Protocols",
        "Communication Network Design",
        "Cybersecurity and Network Programming",
        "Design and Implementation of LAN (CIT 40200)",
        "Digital Forensics",
        "General Information",
        "Human Relations in Organizations",
        "Java Programming I (CIT 27000)",
        "Network Operating System Admin (CIT 35600)",
        "Project Management",
        "Quantitative Analysis",
        "Quantitative Analysis III (32000)",
        "Security Risk Assessment",
        "UNIX Programming and Administration",
        "Wireless Communication",
        "Wireless Security"
    )
)

$OneNote = New-Object -ComObject OneNote.Application

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

# Get full hierarchy including sections
[string]$fullXml = ""
$OneNote.GetHierarchy($notebookId, 4, [ref]$fullXml)  # hsPages = 4, includes sections
[xml]$fullDoc = $fullXml
$nsMgr2 = New-Object System.Xml.XmlNamespaceManager($fullDoc.NameTable)
$nsMgr2.AddNamespace("one", $ns)

New-Item -ItemType Directory -Path $OutputPath -Force | Out-Null

$results = @()

foreach ($sectionName in $SectionNames) {
    # Escape single quotes in XPath if any section name has them
    $escapedName = $sectionName.Replace("'", "''")
    $sectionNode = $fullDoc.SelectSingleNode("//one:Section[@name='$escapedName']", $nsMgr2)

    if (-not $sectionNode) {
        Write-Host "  [not found] $sectionName" -ForegroundColor Red
        $results += [PSCustomObject]@{ Section = $sectionName; Status = "Section not found" }
        continue
    }

    $sectionId = $sectionNode.ID
    $safeFileName = ($sectionName -replace '[\\/:\*\?"<>\|]', '-')
    $destFile = Join-Path $OutputPath "$safeFileName.pdf"

    if (Test-Path $destFile) {
        Write-Host "  [skip - exists] $sectionName" -ForegroundColor Yellow
        $results += [PSCustomObject]@{ Section = $sectionName; Status = "Skipped (already exists)" }
        continue
    }

    try {
        $OneNote.Publish($sectionId, $destFile, 3, "")  # 3 = pfPDF
        Write-Host "  [exported] $sectionName" -ForegroundColor Green
        $results += [PSCustomObject]@{ Section = $sectionName; Status = "Exported"; Path = $destFile }
    } catch {
        Write-Host "  [FAILED] $sectionName - $($_.Exception.Message)" -ForegroundColor Red
        $results += [PSCustomObject]@{ Section = $sectionName; Status = "Failed"; Detail = $_.Exception.Message }
    }
}

$reportPath = Join-Path $OutputPath "_section-export-report.csv"
$results | Export-Csv -Path $reportPath -NoTypeInformation

Write-Host ""
Write-Host "Done. Report: $reportPath" -ForegroundColor Cyan
