# ADO TFVC-to-Git Migrator

A PowerShell toolkit for migrating TFVC repositories from Azure DevOps Server 2022 to Git and GitHub Enterprise. Every script has a guided **interactive mode** with menus, so no command-line experience is required.

## Getting Started (Interactive)

The easiest way to use the toolkit is through the main menu:

```powershell
./Start-Menu.ps1
```

This launches a menu-driven interface:

| Option | Name | What it does |
|--------|------|-------------|
| **1** | Setup Wizard | Walk through server URL, collection PATs, and GitHub settings |
| **2** | Check Prerequisites | Verify git, git-tfs, ImportExcel, and other tools are installed |
| **3** | Discover Repos | Scan all TFVC repos across your collections |
| **4** | Convert Repo | Convert a single TFVC repo to Git |
| **5** | Split Repo | Break one large repo into smaller Git repos by folder |
| **6** | Move Repo | Move a repo to a different ADO collection |
| **7** | Push to GitHub | Send a converted repo to GitHub Enterprise |
| **8** | Run Migration Plan | Execute a saved migration plan JSON file |
| **9** | Batch Migrate | Read the MDR spreadsheet and migrate/split repos in bulk |
| **10** | Batch Archive | Archive repos from the Dalptfs01 spreadsheet |
| **11** | View Logs | Open the logs folder |

No need to remember parameter names or edit JSON by hand.

## Excel-Driven Batch Migration (Option 9)

The primary workflow for McDermott's migration. Reads the **MDR-4ADO-AllProjects.xlsx** spreadsheet ("GAMS-Repos-App-Folder level" worksheet) and processes every row:

- **Column G = "Repo" + Recommendation = "Migrate"** → Converts the entire TFVC repo to Git as-is. If multiple spreadsheet rows reference the same repo (one per subfolder), it's only migrated once.
- **Column G = "Folder" + Recommendation = "Migrate"** → Extracts that folder from the parent repo into a standalone Git repo named `RepoName_FolderName`. All folders from the same parent are extracted in a single batch.
- **Recommendation = "Archive"** → Skipped (not migrated).
- **Collections not in your config** → Silently ignored (only collections you've set up with PATs are processed).

The 4-step interactive flow:

1. **Locate the spreadsheet** — auto-finds `excel-docs/MDR-4ADO-AllProjects.xlsx`
2. **Choose the target** — pick the destination ADO collection and project
3. **Preview** — shows exactly what will be migrated, extracted, archived, and skipped; saves a preview CSV
4. **Confirm and run** — type `yes` to proceed; produces a manifest CSV and JSON report

```powershell
# Interactive (recommended)
./Invoke-ExcelMigration.ps1 -ConfigPath ./config/migration-config.json -Interactive

# Direct
./Invoke-ExcelMigration.ps1 -ConfigPath ./config/migration-config.json `
    -ExcelPath ./excel-docs/MDR-4ADO-AllProjects.xlsx `
    -TargetCollection "ModernApps" -TargetProject "Platform"
```

## Batch Archive (Option 10)

Reads the **Dalptfs01-Collections-MikeFelder.xlsx** spreadsheet and archives repos flagged for decommissioning.

```powershell
./Invoke-ArchiveRepos.ps1 -ConfigPath ./config/migration-config.json -Interactive
```

## Getting Started (Command-Line)

If you prefer scripting or automation, every script also accepts direct parameters:

```powershell
# 1. Create config via the setup wizard (or copy the example)
./New-MigrationConfig.ps1
# — or —
cp config/migration-config.example.json config/migration-config.json

# 2. Verify prerequisites
./Install-Prerequisites.ps1

# 3. Discover what's in your TFVC repos
./Invoke-TfvcDiscovery.ps1 -ConfigPath ./config/migration-config.json

# 4a. Convert an entire TFVC repo to Git
./Convert-TfvcToGit.ps1 -ConfigPath ./config/migration-config.json `
    -Collection "GAMS" -ProjectName "MyProject" -TfvcPath "$/MyProject"

# 4b. Split specific folders into separate Git repos
./Split-TfvcToGitRepos.ps1 -ConfigPath ./config/migration-config.json `
    -Collection "GAMS" -ProjectName "MyProject" -TfvcPath "$/MyProject" `
    -FolderMappings @{
        '$/MyProject/AppA' = 'app-a-repo'
        '$/MyProject/AppB' = 'app-b-repo'
    }

# 5. Push to GitHub Enterprise
./Push-ToGitHub.ps1 -ConfigPath ./config/migration-config.json `
    -RepoPath ./output/app-a-repo -GitHubOrg "McDermott" -GitHubRepo "app-a-repo"
```

## Scripts

| Script | Purpose | Interactive? |
|---|---|---|
| `Start-Menu.ps1` | Main launcher — single entry point for all operations | Yes (always) |
| `New-MigrationConfig.ps1` | Setup wizard — creates `migration-config.json` | Yes (always) |
| `Install-Prerequisites.ps1` | Checks for required tools and auto-installs what it can | — |
| `Invoke-TfvcDiscovery.ps1` | Scans collections and inventories all TFVC repos/folders | `-Interactive` |
| `Convert-TfvcToGit.ps1` | Converts a TFVC repo to a Git repo via git-tfs | `-Interactive` |
| `Split-TfvcToGitRepos.ps1` | Splits TFVC subfolders into separate Git repos | `-Interactive` |
| `Move-RepoToCollection.ps1` | Moves/clones a TFVC repo to a different ADO collection as Git | `-Interactive` |
| `Push-ToGitHub.ps1` | Pushes a converted Git repo to GitHub Enterprise | `-Interactive` |
| `Invoke-ExcelMigration.ps1` | Batch migrate/split repos from the MDR spreadsheet | `-Interactive` |
| `Invoke-ArchiveRepos.ps1` | Batch archive repos from the Dalptfs01 spreadsheet | `-Interactive` |
| `Start-Migration.ps1` | Batch orchestrator — runs a migration plan JSON | — |

## Prerequisites

Run `Install-Prerequisites.ps1` (or pick option **[2]** from the main menu) to verify:

- **PowerShell 7+**
- **git** (2.30+)
- **git-tfs** — the bridge between TFVC and Git ([github.com/git-tfs/git-tfs](https://github.com/git-tfs/git-tfs))
- **git-filter-repo** (optional, for faster folder splitting — `pip install git-filter-repo`)
- **ImportExcel** PowerShell module — needed for the Excel-driven batch scripts (`Install-Module ImportExcel -Scope CurrentUser -Force`)
- **Azure DevOps Server 2022** network connectivity
- A **PAT (Personal Access Token)** per collection with `Code (Read & Write)` and `Project and Team (Read)` scopes

> **Important:** The ADO server URL in your config must use `https://`. git-tfs requires a secure connection for PAT authentication.

## Configuration

You can create the config file in two ways:

1. **Setup wizard** — run `./New-MigrationConfig.ps1` or pick option **[1]** from the main menu. It walks you through each setting and tests your ADO connections.
2. **Manual** — copy [`config/migration-config.example.json`](config/migration-config.example.json) to `config/migration-config.json` and edit it.

Key settings:

| Setting | Description |
|---|---|
| `adoServerUrl` | Your ADO 2022 base URL (e.g. `https://ado.mcdermott.com`) |
| `collections` | Map of collection names, each with a `pat` and optional `description` |
| `outputDirectory` | Where converted Git repos are written (default: `./output`) |
| `logDirectory` | Where log files are written (default: `./logs`) |
| `gitTfsPath` | Path to git-tfs if it's not in your PATH |
| `authorMappingFile` | Optional CSV mapping TFVC usernames → Git authors |
| `github.enterpriseUrl` | GitHub Enterprise URL (e.g. `https://github.mcdermott.com`) |
| `github.pat` | GitHub PAT with `repo` scope |
| `github.defaultOrg` | Default GitHub organization for new repos |

## Directory Structure

```
ado-tfvc-git-migrator/
├── config/
│   ├── migration-config.example.json
│   ├── migration-plan.example.json
│   └── author-mapping.example.csv
├── excel-docs/
│   ├── MDR-4ADO-AllProjects.xlsx      # Batch migration spreadsheet
│   └── Dalptfs01-Collections-MikeFelder.xlsx  # Archive spreadsheet
├── modules/
│   └── AdoTfvcMigrator.psm1          # Shared function library
├── output/                            # Converted repos land here
├── logs/                              # Timestamped migration logs
├── Start-Menu.ps1                     # Main launcher
├── New-MigrationConfig.ps1            # Config setup wizard
├── Install-Prerequisites.ps1          # Prerequisite checker / installer
├── Invoke-TfvcDiscovery.ps1           # Repo discovery & inventory
├── Convert-TfvcToGit.ps1             # Single TFVC-to-Git conversion
├── Split-TfvcToGitRepos.ps1           # Folder-level split
├── Move-RepoToCollection.ps1         # Cross-collection move
├── Push-ToGitHub.ps1                  # GitHub Enterprise push
├── Invoke-ExcelMigration.ps1         # Excel-driven batch migration
├── Invoke-ArchiveRepos.ps1           # Excel-driven batch archive
├── Start-Migration.ps1               # Migration plan orchestrator
└── README.md
```

## Author Mapping

TFVC commits use `DOMAIN\username`. To map these to proper Git authors, create an author mapping CSV:

```csv
TfvcIdentity,GitName,GitEmail
MCDERMOTT\jsmith,John Smith,jsmith@mcdermott.com
MCDERMOTT\jdoe,Jane Doe,jdoe@mcdermott.com
```

Set the path via config (`authorMappingFile`) or generate a template during discovery:

```powershell
# Interactive — the wizard will ask if you want to generate one
./Invoke-TfvcDiscovery.ps1 -ConfigPath ./config/migration-config.json -Interactive

# Direct
./Invoke-TfvcDiscovery.ps1 -ConfigPath ./config/migration-config.json -GenerateAuthorMap
```

## Batch Migrations

For migrating many repos at once, create a migration plan JSON (see [`config/migration-plan.example.json`](config/migration-plan.example.json)) and run:

```powershell
./Start-Migration.ps1 -ConfigPath ./config/migration-config.json `
    -PlanPath ./config/migration-plan.json

# Dry run first to validate the plan without making changes
./Start-Migration.ps1 -ConfigPath ./config/migration-config.json `
    -PlanPath ./config/migration-plan.json -DryRun
```

Or pick option **[8]** from the main menu.

## Logging

All operations write timestamped logs to `./logs/`. View recent logs from the main menu (option **[9]**) or browse the directory directly. Each script also supports `-Verbose` for detailed console output.

## Troubleshooting

Common issues and what to do:

| Symptom | Likely Cause | Fix |
|---|---|---|
| "Authentication failed — your PAT may be expired" | PAT expired or wrong | Generate a new PAT in ADO and update your config |
| "Access denied — your PAT doesn't have the required permissions" | Missing PAT scopes | Ensure `Code (Read & Write)` and `Project and Team (Read)` |
| "Basic authentication requires a secure connection" | Server URL uses `http://` | Change `adoServerUrl` in your config to `https://` |
| "Cannot reach the ADO server" | Network/VPN issue | Check VPN connection and server URL in config |
| "git-tfs not found" | Tool not installed | Run `Install-Prerequisites.ps1` for instructions |
| "Missing Required Module: ImportExcel" | ImportExcel not installed | Run `Install-Module ImportExcel -Scope CurrentUser -Force` |
| Conversion seems frozen | Large repo with lots of history | Normal — a spinner shows elapsed time; check logs for progress |
