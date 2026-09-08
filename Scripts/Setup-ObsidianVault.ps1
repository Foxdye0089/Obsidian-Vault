<#
.SYNOPSIS
    Downloads and installs community plugins directly into the Obsidian vault.

.DESCRIPTION
    Community plugins are just folders of files (main.js, manifest.json,
    optionally styles.css) inside .obsidian/plugins/<plugin-id>/. This script
    pulls the latest GitHub release for each plugin below and drops the files
    in the right place, then enables them in community-plugins.json.

    Safe to re-run — it overwrites each plugin's files with the latest
    release, so you can rerun this later to update everything at once.

.NOTES
    After running this, open Obsidian and go to:
    Settings -> Community plugins -> Turn on community plugins (if not
    already on). Obsidian requires this toggle be confirmed manually the
    first time as a safety measure — it can't be fully scripted.
#>

param(
    [string]$VaultPath = (Join-Path $env:USERPROFILE "Obsidian\Main Vault")
)

# Plugin definitions: id must match the plugin's manifest "id" field exactly
$plugins = @(
    @{ Id = "templater-obsidian";          Repo = "SilentVoid13/Templater" }
    @{ Id = "auto-note-mover";             Repo = "farux/obsidian-auto-note-mover" }
    @{ Id = "dataview";                    Repo = "blacksmithgu/obsidian-dataview" }
    @{ Id = "calendar";                    Repo = "liamcain/obsidian-calendar-plugin" }
    @{ Id = "table-editor-obsidian";       Repo = "tgrosinger/advanced-tables-obsidian" }
    @{ Id = "obsidian-excalidraw-plugin";  Repo = "zsviczian/obsidian-excalidraw-plugin" }
    @{ Id = "obsidian-tasks-plugin";       Repo = "obsidian-tasks-group/obsidian-tasks" }
    @{ Id = "quickadd";                    Repo = "chhoumann/quickadd" }
    @{ Id = "obsidian-kanban";             Repo = "mgmeyers/obsidian-kanban" }
    @{ Id = "periodic-notes";              Repo = "liamcain/obsidian-periodic-notes" }
    @{ Id = "obsidian-git";                Repo = "denolehov/obsidian-git" }
    @{ Id = "nldates-obsidian";            Repo = "argenos/nldates-obsidian" }
    @{ Id = "claudian";                    Repo = "YishenTu/claudian" }
    @{ Id = "pdf-plus";                    Repo = "RyotaUshio/obsidian-pdf-plus" }
    @{ Id = "breadcrumbs";                 Repo = "michaelpporter/breadcrumbs" }
    @{ Id = "homepage";                    Repo = "mirnovov/obsidian-homepage" }
    @{ Id = "todoist";                     Repo = "jamiebrynes7/obsidian-todoist-plugin" }
)

$pluginsRoot = Join-Path $VaultPath ".obsidian\plugins"
New-Item -ItemType Directory -Path $pluginsRoot -Force | Out-Null

Write-Host "Installing plugins into: $pluginsRoot" -ForegroundColor Cyan
Write-Host ""

$installedIds = @()

foreach ($plugin in $plugins) {
    $id   = $plugin.Id
    $repo = $plugin.Repo
    $destFolder = Join-Path $pluginsRoot $id

    Write-Host "Fetching $id (from $repo)..." -ForegroundColor Yellow

    try {
        $releaseUrl = "https://api.github.com/repos/$repo/releases/latest"
        $release = Invoke-RestMethod -Uri $releaseUrl -Headers @{ "User-Agent" = "obsidian-plugin-installer" }

        New-Item -ItemType Directory -Path $destFolder -Force | Out-Null

        # Download whichever of these three files exist in the release assets
        $wanted = @("main.js", "manifest.json", "styles.css")

        foreach ($fileName in $wanted) {
            $asset = $release.assets | Where-Object { $_.name -eq $fileName }
            if ($asset) {
                $destFile = Join-Path $destFolder $fileName
                Invoke-WebRequest -Uri $asset.browser_download_url -OutFile $destFile -UseBasicParsing
                Write-Host "  [downloaded] $fileName" -ForegroundColor Green
            }
            elseif ($fileName -ne "styles.css") {
                # styles.css is optional; main.js and manifest.json are not
                Write-Host "  [missing] $fileName not found in latest release" -ForegroundColor Red
            }
        }

        $installedIds += $id
    }
    catch {
        Write-Host "  [error] Failed to install $id : $_" -ForegroundColor Red
    }

    Write-Host ""
}

# Enable all successfully installed plugins
$communityPluginsPath = Join-Path $VaultPath ".obsidian\community-plugins.json"

$existingIds = @()
if (Test-Path $communityPluginsPath) {
    try {
        $existingIds = Get-Content $communityPluginsPath -Raw | ConvertFrom-Json
    } catch {
        $existingIds = @()
    }
}

$allIds = ($existingIds + $installedIds) | Select-Object -Unique
$allIds | ConvertTo-Json | Set-Content -Path $communityPluginsPath -Encoding UTF8

Write-Host "Enabled plugins list written to community-plugins.json" -ForegroundColor Cyan
Write-Host ""
Write-Host "Done. Installed: $($installedIds -join ', ')" -ForegroundColor White
Write-Host ""
Write-Host "Next: open Obsidian -> Settings -> Community plugins -> confirm 'Turn on community plugins' if prompted." -ForegroundColor Yellow
