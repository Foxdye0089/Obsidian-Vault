<#
.SYNOPSIS
    Cleanly rebuilds a single course's content in the vault: wipes and
    recopies fresh Markdown from the original export, then robustly
    extracts every real attachment for that course and links it directly
    into its matching note.

.DESCRIPTION
    Scoped to one course at a time so results can be verified before
    moving to the next. Places each attachment based on where its note
    ACTUALLY ends up after the fresh copy (not an independently computed
    path), avoiding the folder-mismatch issues from earlier attempts.
    Includes retry-with-reconnect per page so a flaky page can't cause a
    silent skip.

.NOTES
    This DELETES and recopies the course folder in the vault. Only the
    course you specify is touched - nothing else in the vault.
#>

param(
    [Parameter(Mandatory = $true)]
    [string]$NotebookName,

    [Parameter(Mandatory = $true)]
    [string]$CourseName,

    [string]$MarkdownSourcePath = "C:\Users\dakota\Documents\OneNoteExport\Purdue\Purdue",
    [string]$VaultCoursesPath   = "C:\Users\dakota\Obsidian\Main Vault\School\01 - Courses",
    [string]$StagingPath        = "C:\Users\dakota\Backups\Purdue-Attachments",

    # Pages to leave completely alone - backed up before the wipe, restored
    # after, and skipped by attachment recovery. Use the exact page name.
    [string[]]$SkipPages = @()
)

function New-OneNoteConnection {
    return New-Object -ComObject OneNote.Application
}

$script:restartCount = 0
$script:maxRestarts = 2

function Restart-OneNoteProcess {
    $script:restartCount++
    if ($script:restartCount -gt $script:maxRestarts) {
        Write-Host ""
        Write-Host "[STOPPING] OneNote has needed restarting more than $script:maxRestarts times in this run." -ForegroundColor Red
        Write-Host "Repeatedly force-restarting OneNote risks corrupting its local state (this happened before)." -ForegroundColor Red
        Write-Host "Please check that OneNote is healthy, then re-run this script - it will safely skip anything already completed." -ForegroundColor Yellow
        exit 1
    }
    Write-Host "  (restarting OneNote - attempt $script:restartCount of $script:maxRestarts allowed this run)" -ForegroundColor DarkYellow
    Get-Process -Name "ONENOTE" -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue
    Start-Sleep -Seconds 5
    return New-OneNoteConnection
}

# Step 1: Clean rebuild of this course's Markdown in the vault
$vaultCourseFolder = Join-Path $VaultCoursesPath $CourseName
$sourceCourseFolder = Join-Path $MarkdownSourcePath $CourseName

function Get-NormalizedName([string]$name) {
    $normalized = $name -replace '[\\/:\*\?"<>\|\-_]', ' '
    $normalized = $normalized -replace '\s+', ' '
    return $normalized.Trim()
}

# Back up any protected pages BEFORE wiping the course folder, so manual
# fixes survive the rebuild
$protectedBackups = @{}
if ($SkipPages.Count -gt 0 -and (Test-Path $vaultCourseFolder)) {
    Write-Host "Backing up protected pages before wipe..." -ForegroundColor Cyan
    $existingNotes = Get-ChildItem -Path $vaultCourseFolder -Recurse -Filter "*.md" -File -ErrorAction SilentlyContinue
    foreach ($skipName in $SkipPages) {
        $skipNormalized = Get-NormalizedName $skipName
        $match = $existingNotes | Where-Object { (Get-NormalizedName $_.BaseName) -eq $skipNormalized } | Select-Object -First 1
        if ($match) {
            $relativePath = $match.FullName.Substring($vaultCourseFolder.Length).TrimStart('\')
            $protectedBackups[$relativePath] = Get-Content -Path $match.FullName -Raw
            Write-Host "  [protected] $relativePath" -ForegroundColor Yellow
        } else {
            Write-Host "  [warning] Could not find protected page to back up: $skipName" -ForegroundColor Yellow
        }
    }
    Write-Host ""
}

Write-Host "Step 1: Rebuilding '$CourseName' from clean Markdown export..." -ForegroundColor Cyan

if (Test-Path $vaultCourseFolder) {
    Remove-Item -Path $vaultCourseFolder -Recurse -Force
    Write-Host "  Removed old vault folder." -ForegroundColor Yellow
}

robocopy $sourceCourseFolder $vaultCourseFolder /E /R:3 /W:5 /NFL /NDL | Out-Null
Write-Host "  Recopied fresh." -ForegroundColor Green

# Fix duplicate titles: the exporter creates two H1 headers whenever the
# page name has a Windows-illegal character (":", etc.) - one from the
# sanitized filename, one from the real title text. Keep the real one.
Write-Host "  Cleaning up duplicate titles..." -ForegroundColor Cyan

function Get-TitleNormalized([string]$title) {
    $normalized = $title -replace '[\\/:\*\?"<>\|_]', ' '
    $normalized = $normalized -replace '\s+', ' '
    return $normalized.Trim()
}

$allNotes = Get-ChildItem -Path $vaultCourseFolder -Recurse -Filter "*.md" -File
$titleFixCount = 0

foreach ($note in $allNotes) {
    $noteContent = Get-Content -Path $note.FullName -Raw
    $noteContent = $noteContent.TrimStart([char]0xFEFF)  # strip UTF-8 BOM if present - it was breaking the ^# match
    $noteLines = $noteContent -split "`r?`n"

    if ($noteLines.Count -ge 3 -and $noteLines[0] -match '^#\s+(.+)$' -and $noteLines[2] -match '^#\s+(.+)$') {
        $noteLines[0] -match '^#\s+(.+)$' | Out-Null
        $firstTitle = $Matches[1]
        $noteLines[2] -match '^#\s+(.+)$' | Out-Null
        $secondTitle = $Matches[1]

        if ((Get-TitleNormalized $firstTitle) -eq (Get-TitleNormalized $secondTitle) -and $firstTitle -ne $secondTitle) {
            # Keep the second (real) title, drop the first (sanitized) one
            $remainder = $noteLines[3..($noteLines.Length - 1)] -join "`r`n"
            $newContent = "# $secondTitle`r`n`r`n" + $remainder.TrimStart("`r", "`n")
            Set-Content -Path $note.FullName -Value $newContent -Encoding UTF8
            $titleFixCount++
        }
    }
}
Write-Host "  Fixed $titleFixCount duplicate title(s)." -ForegroundColor Green

# Restore protected pages exactly as they were, overwriting whatever the
# fresh copy and title-cleanup just did to them
if ($protectedBackups.Count -gt 0) {
    Write-Host "  Restoring protected pages..." -ForegroundColor Cyan
    foreach ($relativePath in $protectedBackups.Keys) {
        $restorePath = Join-Path $vaultCourseFolder $relativePath
        $restoreDir = Split-Path $restorePath -Parent
        if (-not (Test-Path $restoreDir)) { New-Item -ItemType Directory -Path $restoreDir -Force | Out-Null }
        Set-Content -Path $restorePath -Value $protectedBackups[$relativePath] -Encoding UTF8 -NoNewline
        Write-Host "    [restored] $relativePath" -ForegroundColor Green
    }
}
Write-Host ""

# Step 2: Fresh staging folder for this course's attachments
$stagingCourseFolder = Join-Path $StagingPath $CourseName
if (Test-Path $stagingCourseFolder) {
    Remove-Item -Path $stagingCourseFolder -Recurse -Force
}
New-Item -ItemType Directory -Path $stagingCourseFolder -Force | Out-Null

# Step 3: Connect to OneNote and find this course's pages
$OneNote = New-OneNoteConnection

[string]$notebooksXml = ""
$OneNote.GetHierarchy("", 2, [ref]$notebooksXml)
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

[string]$fullXml = ""
$OneNote.GetHierarchy($notebookId, 4, [ref]$fullXml)
[xml]$fullDoc = $fullXml
$nsMgr2 = New-Object System.Xml.XmlNamespaceManager($fullDoc.NameTable)
$nsMgr2.AddNamespace("one", $ns)

$allPageNodes = $fullDoc.SelectNodes("//one:Page", $nsMgr2)

# Filter to only pages whose top-level ancestor matches our target course
$coursePages = @()
foreach ($pageNode in $allPageNodes) {
    $ancestor = $pageNode.ParentNode
    $topAncestorName = $null
    while ($ancestor -and $ancestor.LocalName -ne "Notebook") {
        if ($ancestor.name) { $topAncestorName = $ancestor.name }
        $ancestor = $ancestor.ParentNode
    }
    if ($topAncestorName -eq $CourseName) {
        $coursePages += $pageNode
    }
}

Write-Host "Step 2: Found $($coursePages.Count) pages in '$CourseName'. Checking for real attachments..." -ForegroundColor Cyan
Write-Host ""

$results = @()

foreach ($pageNode in $coursePages) {
    $pageId = $pageNode.ID
    $pageName = $pageNode.name

    if ($SkipPages -contains $pageName) {
        Write-Host "  [skipped - protected] $pageName" -ForegroundColor DarkGray
        continue
    }

    [string]$pageXml = ""
    $readOk = $false
    for ($attempt = 1; $attempt -le 3; $attempt++) {
        try {
            $OneNote.GetPageContent($pageId, [ref]$pageXml, 0)
            $readOk = $true
            break
        } catch {
            Write-Host "  [retry $attempt] page read failed for '$pageName': $($_.Exception.Message)" -ForegroundColor Red
            $OneNote = Restart-OneNoteProcess
        }
    }

    if (-not $readOk) {
        Write-Host "  [ERROR] Could not read page after 3 attempts: $pageName" -ForegroundColor Red
        $results += [PSCustomObject]@{ Page = $pageName; Status = "Could not read page after retry"; Detail = "" }
        continue
    }

    [xml]$pageDoc = $pageXml
    $nsMgr3 = New-Object System.Xml.XmlNamespaceManager($pageDoc.NameTable)
    $nsMgr3.AddNamespace("one", $ns)

    $insertedFiles = $pageDoc.SelectNodes("//one:InsertedFile", $nsMgr3)
    if ($insertedFiles.Count -eq 0) { continue }

    # Find this page's actual note location in the freshly-copied vault
    # folder. Match by a NORMALIZED name rather than an exact filename,
    # Match by a NORMALIZED name rather than an exact filename, since
    # Windows-illegal characters get sanitized differently between
    # OneNote's page name and the exported filename (":" -> "_", "/" -> "-", etc.)
    # - uses the shared Get-NormalizedName function defined near the top
    $targetNormalized = Get-NormalizedName $pageName
    $allNotesInCourse = Get-ChildItem -Path $vaultCourseFolder -Recurse -Filter "*.md" -File -ErrorAction SilentlyContinue
    $found = $allNotesInCourse | Where-Object { (Get-NormalizedName $_.BaseName) -eq $targetNormalized }

    if ($found.Count -eq 0) {
        Write-Host "  [warning] No note found for page with attachment: $pageName" -ForegroundColor Yellow
        $results += [PSCustomObject]@{ Page = $pageName; Status = "Attachment found but no matching note"; Detail = "" }
        continue
    }
    if ($found.Count -gt 1) {
        Write-Host "  [warning] Multiple notes named '$noteFileName' - ambiguous" -ForegroundColor Yellow
        $results += [PSCustomObject]@{ Page = $pageName; Status = "Ambiguous note match"; Detail = ($found.FullName -join "; ") }
        continue
    }

    $notePath = $found[0].FullName
    $noteFolder = Split-Path $notePath -Parent

    $content = Get-Content -Path $notePath -Raw
    $content = $content.TrimStart([char]0xFEFF)
    $embedLines = @()
    $linkLines = @()

    # Track how many attachments per extension so same-page duplicates
    # (e.g. two PDFs) don't collide on the same renamed filename
    $extensionCounts = @{}

    foreach ($file in $insertedFiles) {
        $pathCache = $file.pathCache
        $preferredName = $file.preferredName

        if ([string]::IsNullOrWhiteSpace($pathCache) -or -not (Test-Path $pathCache)) {
            Write-Host "  [warning] Cached file missing on disk for '$pageName' -> $preferredName (path: $pathCache)" -ForegroundColor Yellow
            $results += [PSCustomObject]@{ Page = $pageName; Status = "No recoverable cached copy"; Detail = $preferredName }
            continue
        }

        # Copy the original, messily-named file to staging as-is (backup
        # just needs to preserve the real files, naming doesn't matter there)
        $stagingDest = Join-Path $stagingCourseFolder $preferredName
        Copy-Item -Path $pathCache -Destination $stagingDest -Force

        # For the vault copy: rename to match the page name, into a nested
        # Attachments folder to keep the note list uncluttered. Taking just
        # the FINAL extension naturally collapses double-extension names
        # like "file.pptx.pdf" down to a clean ".pdf"
        $ext = [System.IO.Path]::GetExtension($preferredName)
        $safePageName = ($pageName -replace '[\\/:\*\?"<>\|]', '-').Trim()

        $extKey = $ext.ToLower()
        if (-not $extensionCounts.ContainsKey($extKey)) { $extensionCounts[$extKey] = 0 }
        $extensionCounts[$extKey]++
        $suffix = if ($extensionCounts[$extKey] -gt 1) { " ($($extensionCounts[$extKey]))" } else { "" }

        $renamedFile = "$safePageName$suffix$ext"

        $attachmentsFolder = Join-Path $noteFolder "Attachments"
        if (-not (Test-Path $attachmentsFolder)) {
            New-Item -ItemType Directory -Path $attachmentsFolder -Force | Out-Null
        }
        $vaultDest = Join-Path $attachmentsFolder $renamedFile
        Copy-Item -Path $pathCache -Destination $vaultDest -Force

        if ($extKey -eq ".pdf") {
            $embedLines += "![[$renamedFile]]"
        } else {
            $linkLines += "Original file: [[$renamedFile]]"
        }

        Write-Host "  [recovered] $pageName -> $renamedFile" -ForegroundColor Green
        $results += [PSCustomObject]@{ Page = $pageName; Status = "Recovered and linked"; Detail = $renamedFile }
    }

    if ($embedLines.Count -gt 0 -or $linkLines.Count -gt 0) {
        # Strip any broken-image junk text, then insert clean links/embeds
        $cleaned = [regex]::Replace($content, '\*?\[Image - no embedded data.*?\]\*?\r?\n?', '')
        $cleaned = [regex]::Replace($cleaned, '(\r?\n){3,}', "`r`n`r`n")

        $insertBlock = $linkLines + @("") + $embedLines

        $lines = $cleaned -split "`r?`n"
        if ($lines[0] -match "^#\s") {
            $newLines = @($lines[0], "") + $insertBlock + @("") + $lines[1..($lines.Length - 1)]
        } else {
            $newLines = $insertBlock + @("") + $lines
        }

        ($newLines -join "`r`n") | Set-Content -Path $notePath -Encoding UTF8
    }
}

$reportPath = Join-Path $stagingCourseFolder "_recovery-report.csv"
$results | Export-Csv -Path $reportPath -NoTypeInformation

Write-Host ""
$recoveredCount = ($results | Where-Object { $_.Status -eq "Recovered and linked" }).Count
Write-Host "Done. Recovered and linked $recoveredCount attachment(s) for '$CourseName'." -ForegroundColor Cyan

# Step 4: Whole-section PDF as a single-file safety net, on top of the
# individual attachments above. If OneNote/the account ever becomes
# fully inaccessible, this is the fallback that still has everything.
Write-Host ""
Write-Host "Step 3: Exporting whole-section PDF as a backup safety net..." -ForegroundColor Cyan

$escapedCourseName = $CourseName.Replace("'", "''")
$sectionNode = $fullDoc.SelectSingleNode("//one:Section[@name='$escapedCourseName']", $nsMgr2)

if (-not $sectionNode) {
    Write-Host "  [warning] Could not find section node for '$CourseName' - skipping PDF backup." -ForegroundColor Yellow
} else {
    $sectionId = $sectionNode.ID
    $sectionPdfBackup = Join-Path $stagingCourseFolder "$CourseName - Full Section Export.pdf"
    $sectionPdfVault = Join-Path $vaultCourseFolder "$CourseName - Full Section Export.pdf"

    $exported = $false
    for ($attempt = 1; $attempt -le 3; $attempt++) {
        try {
            $OneNote.Publish($sectionId, $sectionPdfBackup, 3, "")
            $exported = $true
            break
        } catch {
            Write-Host "  attempt $attempt failed: $($_.Exception.Message)" -ForegroundColor Red
            $OneNote = Restart-OneNoteProcess
        }
    }

    if ($exported) {
        Copy-Item -Path $sectionPdfBackup -Destination $sectionPdfVault -Force
        Write-Host "  [done] Section PDF backed up and placed in vault." -ForegroundColor Green
    } else {
        Write-Host "  [FAILED] Could not export section PDF after retry - individual attachments above are still recovered." -ForegroundColor Red
    }
}

Write-Host ""
Write-Host "Staging backup: $stagingCourseFolder" -ForegroundColor Yellow
Write-Host "Report: $reportPath" -ForegroundColor Yellow
