<#
.SYNOPSIS
    Runs Rebuild-CourseAttachments.ps1 across only the courses listed
    below - explicitly, not "everything except the done ones" - so
    there's no risk of accidentally touching a course you've already
    verified is good.

.DESCRIPTION
    Each course still runs independently: if one course throws an
    unexpected error, the batch logs it and moves on to the next rather
    than aborting the whole run.

.NOTES
    Rebuild-CourseAttachments.ps1 must be in the same folder as this
    script (or update $scriptPath below).
#>

param(
    [string]$NotebookName = "IUPUI"
)

# EXPLICIT list - only courses NOT already verified good. Digital Forensics,
# Wireless Communication, UNIX Programming and Administration,
# Wireless Security, Communication Network Design, and Security Risk
# Assessment are deliberately NOT in this list.
$coursesToProcess = @(
    "Adv Network Administration (CIT 41500)",
    "Advanced Network Security (CIT 40600)",
    "Applied Secure Protocols",
    "Cybersecurity and Network Programming",
    "Design and Implementation of LAN (CIT 40200)",
    "General Information",
    "Human Relations in Organizations",
    "Java Programming I (CIT 27000)",
    "Network Operating System Admin (CIT 35600)",
    "Project Management",
    "Quantitative Analysis",
    "Quantitative Analysis III (32000)"
)

$scriptPath = Join-Path $PSScriptRoot "Rebuild-CourseAttachments.ps1"

if (-not (Test-Path $scriptPath)) {
    Write-Host "Could not find Rebuild-CourseAttachments.ps1 in $PSScriptRoot" -ForegroundColor Red
    exit 1
}

Write-Host "About to process $($coursesToProcess.Count) courses:" -ForegroundColor Cyan
$coursesToProcess | ForEach-Object { Write-Host "  - $_" }
Write-Host ""

$batchResults = @()

foreach ($course in $coursesToProcess) {
    Write-Host ""
    Write-Host "==================================================" -ForegroundColor Magenta
    Write-Host "Starting: $course" -ForegroundColor Magenta
    Write-Host "==================================================" -ForegroundColor Magenta

    try {
        & $scriptPath -NotebookName $NotebookName -CourseName $course
        $batchResults += [PSCustomObject]@{ Course = $course; Status = "Completed" }
    } catch {
        Write-Host "[BATCH ERROR] '$course' threw an unexpected error: $($_.Exception.Message)" -ForegroundColor Red
        Write-Host "Moving on to the next course." -ForegroundColor Yellow
        $batchResults += [PSCustomObject]@{ Course = $course; Status = "Failed"; Detail = $_.Exception.Message }
    }
}

Write-Host ""
Write-Host "==================================================" -ForegroundColor Cyan
Write-Host "Batch complete." -ForegroundColor Cyan
$batchResults | Format-Table -AutoSize

$batchReportPath = "C:\Users\dakota\Backups\Purdue-Attachments\_batch-run-report.csv"
$batchResults | Export-Csv -Path $batchReportPath -NoTypeInformation
Write-Host "Batch report: $batchReportPath" -ForegroundColor Yellow
