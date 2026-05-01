# ADO TFVC-to-Git Migrator

A PowerShell toolkit for migrating TFVC repositories from Azure DevOps Server 2022 to Git and GitHub Enterprise. Every script has a guided **interactive mode** with menus, so no command-line experience is required.

## Getting Started (Interactive)

The easiest way to use the toolkit is through the main menu:

```powershell
./Start-Menu.ps1
```

This launches a menu-driven interface where you can:

1. **Set up your configuration** — walk through a wizard that asks for your ADO server URL, collection PATs, and GitHub settings
2. **Check prerequisites** — verify git, git-tfs, and other tools are installed
3. **Discover repos** — scan all TFVC repos across your collections
4. **Convert, split, or move repos** — each operation has a step-by-step wizard
5. **Push to GitHub** — browse your converted repos and push them up

No need to remember parameter names or edit JSON by hand.

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
| `Install-Prerequisites.ps1` | Checks for required tools (PS7, git, git-tfs, etc.) | — |
| `Invoke-TfvcDiscovery.ps1` | Scans collections and inventories all TFVC repos/folders | `-Interactive` |
| `Convert-TfvcToGit.ps1` | Converts a TFVC repo to a Git repo via git-tfs | `-Interactive` |
| `Split-TfvcToGitRepos.ps1` | Splits TFVC subfolders into separate Git repos | `-Interactive` |
| `Move-RepoToCollection.ps1` | Moves/clones a TFVC repo to a different ADO collection as Git | `-Interactive` |
| `Push-ToGitHub.ps1` | Pushes a converted Git repo to GitHub Enterprise | `-Interactive` |
| `Start-Migration.ps1` | Batch orchestrator — runs a migration plan JSON | — |

## Prerequisites

Run `Install-Prerequisites.ps1` (or pick option **[2]** from the main menu) to verify:

- **PowerShell 7+**
- **git** (2.30+)
- **git-tfs** — the bridge between TFVC and Git ([github.com/git-tfs/git-tfs](https://github.com/git-tfs/git-tfs))
- **git-filter-repo** (optional, for faster folder splitting — `pip install git-filter-repo`)
- **Azure DevOps Server 2022** network connectivity
- A **PAT (Personal Access Token)** per collection with `Code (Read & Write)` and `Project and Team (Read)` scopes

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
├── modules/
│   └── AdoTfvcMigrator.psm1          # Shared function library
├── output/                            # Converted repos land here
├── logs/                              # Timestamped migration logs
├── Start-Menu.ps1                     # Main launcher
├── New-MigrationConfig.ps1            # Config setup wizard
├── Install-Prerequisites.ps1
├── Invoke-TfvcDiscovery.ps1
├── Convert-TfvcToGit.ps1
├── Split-TfvcToGitRepos.ps1
├── Move-RepoToCollection.ps1
├── Push-ToGitHub.ps1
├── Start-Migration.ps1               # Batch orchestrator
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
| "Cannot reach the ADO server" | Network/VPN issue | Check VPN connection and server URL in config |
| "git-tfs not found" | Tool not installed | Run `Install-Prerequisites.ps1` for instructions |
| Conversion seems frozen | Large repo with lots of history | Normal — a spinner shows elapsed time; check logs for progress |
