# ADO TFVC-to-Git Migrator

A PowerShell-based toolkit for migrating TFVC repositories from Azure DevOps Server 2022 to Git, supporting the push to GitHub Enterprise.

## Scenarios Supported

| Scenario | Script |
|---|---|
| Inventory all TFVC repos/folders across collections | `Invoke-TfvcDiscovery.ps1` |
| Convert an entire TFVC repo to a Git repo | `Convert-TfvcToGit.ps1` |
| Split TFVC subfolders into separate Git repos | `Split-TfvcToGitRepos.ps1` |
| Move/clone a repo to a different ADO Collection/Project | `Move-RepoToCollection.ps1` |
| Push converted Git repos to GitHub Enterprise | `Push-ToGitHub.ps1` |
| Full orchestrated migration (discover → convert → push) | `Start-Migration.ps1` |

## Prerequisites

Run `Install-Prerequisites.ps1` to verify and install required tools:

- **PowerShell 7+**
- **git** (2.30+)
- **git-tfs** — the primary bridge between TFVC and Git ([github.com/git-tfs/git-tfs](https://github.com/git-tfs/git-tfs))
- **Azure DevOps Server 2022** connectivity (HTTP/HTTPS to the instance)
- A **PAT (Personal Access Token)** with `Code (Read & Write)` and `Project (Read)` scopes on each collection

## Quick Start

```powershell
# 1. Copy and edit the config
cp config/migration-config.example.json config/migration-config.json
# Edit with your ADO server URL, collections, PAT, etc.

# 2. Verify prerequisites
./Install-Prerequisites.ps1

# 3. Discover what's in your TFVC repos
./Invoke-TfvcDiscovery.ps1 -ConfigPath ./config/migration-config.json

# 4a. Convert an entire TFVC repo to Git
./Convert-TfvcToGit.ps1 -ConfigPath ./config/migration-config.json `
    -Collection "GAMS" -ProjectName "MyProject" -TfvcPath "$/MyProject"

# 4b. Split specific folders into separate Git repos
./Split-TfvcToGitRepos.ps1 -ConfigPath ./config/migration-config.json `
    -Collection "GAMS" -ProjectName "MyProject" `
    -FolderMappings @{
        '$/MyProject/AppA' = 'app-a-repo'
        '$/MyProject/AppB' = 'app-b-repo'
    }

# 5. Push to GitHub Enterprise
./Push-ToGitHub.ps1 -RepoPath ./output/app-a-repo `
    -GitHubOrg "McDermott" -GitHubRepo "app-a-repo"
```

## Configuration

See [`config/migration-config.example.json`](config/migration-config.example.json) for all options.

Key settings:
- `adoServerUrl` — Your ADO 2022 base URL (e.g. `https://ado.mcdermott.com`)
- `collections` — Map of collection names to PATs
- `outputDirectory` — Where converted Git repos are written
- `gitTfsPath` — Path to git-tfs executable (if not in PATH)
- `authorMappingFile` — Optional CSV mapping TFVC users → Git authors

## Directory Structure

```
ado-tfvc-git-migrator/
├── config/
│   ├── migration-config.example.json
│   └── author-mapping.example.csv
├── modules/
│   └── AdoTfvcMigrator.psm1        # Shared functions
├── output/                          # Converted repos land here
├── logs/                            # Migration logs
├── Convert-TfvcToGit.ps1
├── Install-Prerequisites.ps1
├── Invoke-TfvcDiscovery.ps1
├── Move-RepoToCollection.ps1
├── Push-ToGitHub.ps1
├── Split-TfvcToGitRepos.ps1
├── Start-Migration.ps1
└── README.md
```

## Author Mapping

TFVC commits use `DOMAIN\username`. To map these to proper Git authors, create an author mapping CSV:

```csv
TfvcIdentity,GitName,GitEmail
MCDERMOTT\jsmith,John Smith,jsmith@mcdermott.com
MCDERMOTT\jdoe,Jane Doe,jdoe@mcdermott.com
```

Pass it via config (`authorMappingFile`) or generate a template with:

```powershell
./Invoke-TfvcDiscovery.ps1 -ConfigPath ./config/migration-config.json -GenerateAuthorMap
```

## Logging

All operations write timestamped logs to `./logs/`. Each script also supports `-Verbose` for detailed console output.
