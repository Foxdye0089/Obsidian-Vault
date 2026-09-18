<#
.SYNOPSIS
    Merges the three recovered Purdue content sources (Markdown pages,
    recovered attachments, section PDFs) into the vault's School folder,
    organized by course.

.DESCRIPTION
    This COPIES (does not move) from each staging location into
    School\01 - Courses\<Course Name>\ in the vault, mirroring the
    original section/page folder structure. Staging folders are left
    untouched so nothing is at risk until you've verified everything
    looks right in Obsidian - clean those up manually once confirmed.

.NOTES
    Edit the three $...Path variables below if your staging folders are
    somewhere other than the defaults used earlier in this process.
#>

param(
    [string]$MarkdownPath    = "C:\Users\dakota\Documents\OneNoteExport\Purdue\Purdue",
    [string]$AttachmentsPath = "C:\Users\dakota\Documents\OneNoteExport\Purdue-Attachments",
    [string]$SectionPdfPath  = "C:\Users\dakota\Documents\OneNoteExport\Purdue-SectionPDFs",
    [string]$VaultCoursesPath = "C:\Users\dakota\Obsidian\Main Vault\School\01 - Courses"
)

New-Item -ItemType Directory -Path $VaultCoursesPath -Force | Out-Null

Write-Host "Step 1: Copying Markdown pages into vault..." -ForegroundColor Cyan
robocopy $MarkdownPath $VaultCoursesPath /E /R:3 /W:5 /LOG+:"$VaultCoursesPath\_merge-log.txt" /NFL /NDL | Out-Null
Write-Host "  Done." -ForegroundColor Green

Write-Host "Step 2: Merging recovered attachments into vault..." -ForegroundColor Cyan
robocopy $AttachmentsPath $VaultCoursesPath /E /R:3 /W:5 /LOG+:"$VaultCoursesPath\_merge-log.txt" /NFL /NDL | Out-Null
Write-Host "  Done." -ForegroundColor Green

Write-Host "Step 3: Placing section PDFs into their matching course folders..." -ForegroundColor Cyan
$pdfFiles = Get-ChildItem -Path $SectionPdfPath -Filter "*.pdf"

foreach ($pdf in $pdfFiles) {
    $courseName = $pdf.BaseName
    $destFolder = Join-Path $VaultCoursesPath $courseName

    if (-not (Test-Path $destFolder)) {
        Write-Host "  [warning] No matching course folder for: $courseName" -ForegroundColor Yellow
        Write-Host "            Creating folder and placing PDF anyway." -ForegroundColor Yellow
        New-Item -ItemType Directory -Path $destFolder -Force | Out-Null
    }

    $destFile = Join-Path $destFolder "$courseName - Full Section Export.pdf"
    Copy-Item -Path $pdf.FullName -Destination $destFile -Force
    Write-Host "  [placed] $courseName -> Full Section Export.pdf" -ForegroundColor Green
}

Write-Host ""
Write-Host "All done. Nothing was deleted from staging folders - verify the vault, then clean those up manually when ready." -ForegroundColor Cyan
Write-Host "Merge log: $VaultCoursesPath\_merge-log.txt" -ForegroundColor Yellow
