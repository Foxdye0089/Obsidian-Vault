# Obsidian Vault Setup

PowerShell scripts to bootstrap a fresh Obsidian vault from scratch — folder
structure, then community plugins — without clicking through everything by
hand. Built for a personal "Main Vault" layout combining Work, Personal,
School, Salesforce, and reference notebooks, but the folder list and plugin
list are easy to edit for a different setup.

## Scripts

### `scripts/Setup-ObsidianVault.ps1`

Creates the top-level `Obsidian/` sync folder and the `Main Vault/` folder
skeleton inside it (Work, Personal, School, Salesforce, Cyberkelp, Templates,
Archive, each with their planned subfolders). Safe to re-run — existing
folders are left alone.

```powershell
.\scripts\Setup-ObsidianVault.ps1
```

Defaults to `%USERPROFILE%\Obsidian\Main Vault`. Pass `-VaultPath` on the
plugin script (see below) if you move it elsewhere.

### `scripts/Install-ObsidianPlugins.ps1`

Downloads the latest GitHub release of each plugin below directly into
`.obsidian/plugins/`, and enables them in `community-plugins.json`. Safe to
re-run later to update everything at once.

```powershell
.\scripts\Install-ObsidianPlugins.ps1
```

| Plugin | Purpose |
|---|---|
| Templater | Scriptable note templates |
| Auto Note Mover | Auto-file tagged notes out of Inbox folders |
| Dataview | Query notes like a database |
| Calendar | Calendar view + daily notes |
| Advanced Tables | Better Markdown table editing |
| Excalidraw | Freeform drawing/diagramming in notes |
| Tasks | Task management across notes |
| QuickAdd | Fast capture via hotkey + template |
| Kanban | Board view for projects |
| Periodic Notes | Scheduled daily/weekly/monthly notes |
| Obsidian Git | Auto-commit vault history to a Git repo |
| Natural Language Dates | Parse dates like "next monday" |
| Claudian | Embeds Claude Code as an AI agent in the vault |
| PDF++ | Native PDF annotation and viewing |
| Breadcrumbs | Typed links / hierarchy navigation |
| Homepage | Open a specific note/dashboard on startup |
| Todoist | Sync Todoist tasks into notes |

Obsidian requires manually confirming **Settings → Community plugins → Turn
on community plugins** the first time — this is a safety gate that can't be
scripted.

## First-time setup on a new machine

Scripts downloaded from the internet (including from a browser or chat tool)
are flagged by Windows and blocked from running until unblocked:

```powershell
Unblock-File -Path .\scripts\Setup-ObsidianVault.ps1
Unblock-File -Path .\scripts\Install-ObsidianPlugins.ps1
```

You may also need to allow script execution for your user account once:

```powershell
Set-ExecutionPolicy -Scope CurrentUser RemoteSigned
```

Then run the two scripts in order: vault skeleton first, plugins second.

## Notes on specific plugins

- **Claudian** requires Node.js and the Claude Code CLI installed and
  configured separately — the plugin files alone won't do anything without
  those. It also sends vault content to whichever AI provider you configure.
- **Todoist** needs an API token pasted into its settings after install
  (Todoist → Settings → Integrations → Developer).
- **Obsidian Git** needs Git installed and the vault initialized as a Git
  repo to actually commit anything.
