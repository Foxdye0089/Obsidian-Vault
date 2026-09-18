<#
.SYNOPSIS
    Scans an exported OneNote-to-Markdown folder and flags pages that are
    suspiciously short — likely pages where the real content was an
    attachment printout (embedded PowerPoint/PDF slide image) that didn't
    make it into the Markdown conversion.

.DESCRIPTION
    A normal notes page has real body text. A page whose only content was
    an inserted file (like a professor's slide deck) often exports as just
    a title with nothing else. This script lists any .md file under the
    word-count threshold so you can manually review just those, instead of
    checking all 244+ pages by hand.

.NOTES
    This is a heuristic, not a guarantee — a genuinely short note (like a
    one-line reminder) will also get flagged. Skim the flagged list; it's
    meant to narrow 244 pages down to a manageable review pile, not to be
    perfectly precise.
#>

param(
    [Parameter(Mandatory = $true)]
    [string]$ExportPath,

    [int]$WordThreshold = 15
)

if (-not (Test-Path $ExportPath)) {
    Write-Host "Path not found: $ExportPath" -ForegroundColor Red
    exit 1
}

$mdFiles = Get-ChildItem -Path $ExportPath -Filter "*.md" -Recurse

$flaggedThin = @()
$flaggedBrokenImage = @()

foreach ($file in $mdFiles) {
    $content = Get-Content -Path $file.FullName -Raw

    # Primary signal: the exporter's own failure message, left inline
    # when it couldn't pull image binary data (attachment printouts,
    # mainly). This is a precise match, not a guess.
    if ($content -match "no embedded data|could not fetch binary content") {
        $imageErrorCount = ([regex]::Matches($content, "no embedded data")).Count
        $flaggedBrokenImage += [PSCustomObject]@{
            Page             = $file.Name
            BrokenImageCount = $imageErrorCount
            Path             = $file.FullName
        }
        continue
    }

    # Secondary signal: suspiciously short body content, in case a page
    # dropped its real content without leaving an error message behind
    $bodyOnly = ($content -split "`n" | Where-Object { $_ -notmatch "^\s*#" }) -join " "
    $wordCount = ($bodyOnly -split '\s+' | Where-Object { $_.Trim() -ne "" }).Count

    if ($wordCount -le $WordThreshold) {
        $flaggedThin += [PSCustomObject]@{
            Page      = $file.Name
            WordCount = $wordCount
            Path      = $file.FullName
        }
    }
}

Write-Host ""
Write-Host "Scanned $($mdFiles.Count) pages." -ForegroundColor Cyan
Write-Host ""
Write-Host "Pages with broken image errors baked in: $($flaggedBrokenImage.Count)" -ForegroundColor Red
if ($flaggedBrokenImage.Count -gt 0) {
    $flaggedBrokenImage | Format-Table -AutoSize
}

Write-Host ""
Write-Host "Pages suspiciously thin (<= $WordThreshold words, no error text): $($flaggedThin.Count)" -ForegroundColor Yellow
if ($flaggedThin.Count -gt 0) {
    $flaggedThin | Format-Table -AutoSize
}

$reportPath = Join-Path $ExportPath "_export-issues-report.csv"
($flaggedBrokenImage + $flaggedThin) | Export-Csv -Path $reportPath -NoTypeInformation
Write-Host ""
Write-Host "Full combined list saved to: $reportPath" -ForegroundColor Cyan
