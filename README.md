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
| **12** | Mirror Git Collection | Mirror every Git repo from one collection/org to another (cross-server supported) |
| **13** | Migrate ONE Project | Interactive picker — see per-project migration status, pick one, mirror it to `MDR-GAMS-ADO` |

No need to remember parameter names or edit JSON by hand.

## Excel-Driven Batch Migration (Option 9)

The primary workflow for McDermott's migration. Reads the **MDR-4ADO-AllProjects.xlsx** spreadsheet ("GAMS-Repos-App-Folder level" worksheet) and processes every row:

- **Column G = "Repo" + Recommendation = "Migrate" + Repos Type = "TFVC"** → Converts the entire TFVC repo to Git as-is. If multiple spreadsheet rows reference the same repo (one per subfolder), it's only migrated once.
- **Column G = "Repo" + Recommendation = "Migrate" + Repos Type = "Git"** → Mirrors the existing Git repo (`git clone --mirror` + `git push --mirror`) from the source collection straight into the **Git target collection** (default `GAMS-GIT-Repos`). Target projects are auto-created if missing; the source project name is preserved unless you pass `-GitTargetProject`. This is the second-pass mode for the rows that were previously skipped because they were already Git.
- **Column G = "Folder" + Recommendation = "Migrate"** → Extracts that folder from the parent repo into a standalone Git repo named `RepoName_FolderName`. All folders from the same parent are extracted in a single batch. Git-source folder rows are skipped with a clear reason (sub-tree extraction from an existing Git repo is out of scope — clone and split it manually if you need that).
- **Recommendation = "Archive"** → Skipped (not migrated).
- **Collections not in your config** → Silently ignored (only collections you've set up with PATs are processed).

The 4-step interactive flow:

1. **Locate the spreadsheet** — auto-finds `excel-docs/MDR-4ADO-AllProjects.xlsx`
2. **Choose the target** — pick the destination ADO collection and project
3. **Preview** — shows exactly what will be migrated, extracted, archived, and skipped; saves a preview CSV
4. **Confirm and run** — type `yes` to proceed; produces a manifest CSV and JSON report

During batch execution, the migration runner now also:

- Runs conversion in non-interactive mode so existing output folders are auto-cleaned (no blocking Y/N prompt)
- Shows per-item progress as `[X / Y] (pct%) | ETA` and logs per-item duration
- Applies timeout and stall detection to git-tfs operations (`timeoutMinutes` and `stallTimeoutMinutes`)
- Marks successful rows in the spreadsheet Recommendation column as `completed`
- Marks stuck items as `TimedOut`, logs them, and continues with the rest of the batch

```powershell
# Interactive (recommended) — prompts for TFVC target and (if Git rows exist) Git target
./Invoke-ExcelMigration.ps1 -ConfigPath ./config/migration-config.json -Interactive

# Direct — TFVC rows only
./Invoke-ExcelMigration.ps1 -ConfigPath ./config/migration-config.json `
    -ExcelPath ./excel-docs/MDR-4ADO-AllProjects.xlsx `
    -TargetCollection "ModernApps" -TargetProject "Platform"

# Direct — second-pass run for Git-source rows, mirrored into GAMS-GIT-Repos
./Invoke-ExcelMigration.ps1 -ConfigPath ./config/migration-config.json `
    -ExcelPath ./excel-docs/MDR-4ADO-AllProjects.xlsx `
    -GitTargetCollection "GAMS-GIT-Repos"

# Direct — process both kinds in one pass with explicit timeout controls
./Invoke-ExcelMigration.ps1 -ConfigPath ./config/migration-config.json `
    -ExcelPath ./excel-docs/MDR-4ADO-AllProjects.xlsx `
    -TargetCollection "ModernApps" -TargetProject "Platform" `
    -GitTargetCollection "GAMS-GIT-Repos" `
    -TimeoutMinutes 120 -StallTimeoutMinutes 30
```

> **Note:** The Git target collection must already exist in your config. The PAT on that entry needs **Code (Read & Write)** and **Project and Team (Read, Write & Manage)** so the script can auto-create missing target projects and Git repos. Each Git row is marked `completed` in column H of the spreadsheet on success, exactly like the TFVC path.

## Batch Archive (Option 10)

Reads the **Dalptfs01-Collections-MikeFelder.xlsx** spreadsheet and archives repos flagged for decommissioning.

```powershell
./Invoke-ArchiveRepos.ps1 -ConfigPath ./config/migration-config.json -Interactive
```

## Mirror Git Collection — Cross-Server / Org (Option 12)

Mirrors **every project and Git repository** from a source ADO collection to a target collection, preserving names. The flagship use case is the on-prem ADO Server 2022 collection `GAMS-GIT-Repos` mirrored to the Azure DevOps Services organisation `MDR-GAMS-ADO`, but it works for any source/target defined in the config.

For each source project:

1. The matching target project is **auto-created** if it doesn't already exist (process template defaults to **Agile**).
2. Every Git repo in the source project is mirrored to the target project of the same name using `git clone --mirror` followed by `git push --mirror`. **All branches, tags, and notes** come across.
3. Existing target repos receive a force-push of all refs.

### Configuration for a cross-server target

Add the target organisation as another collection entry, with an explicit `serverUrl` pointing at Azure DevOps Services. The collection key becomes the org name:

```jsonc
{
    "adoServerUrl": "https://ado.mcdermott.com",
    "collections": {
        "GAMS-GIT-Repos": {
            "pat": "<source on-prem PAT>",
            "description": "Source — on-prem ADO 2022 collection"
        },
        "MDR-GAMS-ADO": {
            "pat": "<target ADO Services PAT>",
            "serverUrl": "https://dev.azure.com",
            "description": "Target — Azure DevOps Services organisation"
        }
    }
}
```

`serverUrl` is optional per collection. When set it overrides the top-level `adoServerUrl` for that one collection — enabling cross-server mirrors without disturbing any other workflow.

### PAT scopes required on the target

The **target** PAT must include both:

- **Code** — Read & Write (push refs into existing/new repos)
- **Project and Team** — Read, Write & Manage (auto-create missing target projects)

The source PAT only needs **Code (Read)**.

### Run it

```powershell
# Interactive — pick source/target, preview, then confirm
./Mirror-AdoCollection.ps1 -ConfigPath ./config/migration-config.json -Interactive

# Direct — on-prem collection -> ADO Services org
./Mirror-AdoCollection.ps1 -ConfigPath ./config/migration-config.json `
    -SourceCollection GAMS-GIT-Repos -TargetCollection MDR-GAMS-ADO -Force

# Preview-only — just write the manifest CSV, do nothing else
./Mirror-AdoCollection.ps1 -ConfigPath ./config/migration-config.json `
    -SourceCollection GAMS-GIT-Repos -TargetCollection MDR-GAMS-ADO -PreviewOnly

# Dry run — walk the whole flow, ensure target projects exist, but skip clones/pushes
./Mirror-AdoCollection.ps1 -ConfigPath ./config/migration-config.json `
    -SourceCollection GAMS-GIT-Repos -TargetCollection MDR-GAMS-ADO -DryRun -Force

# Filter to a subset of source projects
./Mirror-AdoCollection.ps1 -ConfigPath ./config/migration-config.json `
    -SourceCollection GAMS-GIT-Repos -TargetCollection MDR-GAMS-ADO `
    -IncludeProjects 'AppA','AppB' -Force

# Filter to a single source project (also used by option [13])
./Mirror-AdoCollection.ps1 -ConfigPath ./config/migration-config.json `
    -SourceCollection GAMS-GIT-Repos -TargetCollection MDR-GAMS-ADO `
    -SourceProject 'AppA' -Force

# Resume after a partial run — repos with Status=Success in the manifest are skipped
./Mirror-AdoCollection.ps1 -ConfigPath ./config/migration-config.json `
    -SourceCollection GAMS-GIT-Repos -TargetCollection MDR-GAMS-ADO `
    -ResumeManifest ./output/mirror-MANIFEST-20260520-093011.csv -Force
```

### Outputs

Every run writes three files under `outputDirectory`:

- `mirror-PREVIEW-<timestamp>.csv` — list of every source repo to be mirrored
- `mirror-MANIFEST-<timestamp>.csv` — per-repo result (`Success`/`Skipped`/`Failed`/`PathTooLong`/`TimedOut`/`DryRun`) with duration in seconds; flushed after every repo so a crash never loses progress
- `mirror-REPORT-<timestamp>.json` — final summary with totals and file paths

Temporary bare clones live under `./output/mirror-cache/` (override with `-WorkingDirectory`). They're deleted on success unless you pass `-KeepCache`.

## Migrate ONE Project — Interactive Picker (Option 13)

When you only want to mirror a **single project** at a time (instead of the whole collection), pick option **[13]** from the menu — or run [`Invoke-ProjectMigration.ps1`](Invoke-ProjectMigration.ps1) directly. It shows a live, color-coded table of every project in the source collection and its current migration status against the target (`MDR-GAMS-ADO` by default):

| Symbol | Color | Meaning |
|---|---|---|
| `[OK]` | Green | Migrated — every source repo exists in the target project |
| `[~]`  | Yellow | Partial — target exists but is missing *N* repos |
| `[ ]`  | Red | Not migrated — target project doesn't exist (or has none of the source repos) |
| `[--]` | Gray | Empty — source project has no Git repos to mirror |

Each row also shows the **Src / Tgt** repo counts so you can see at a glance how much work remains.

### Keyboard controls

| Key | Action |
|---|---|
| `↑` / `↓` / `PageUp` / `PageDown` / `Home` / `End` | Move the cursor |
| `1`–`9`… then `Enter` | Jump to a project by its row number (numbering is stable even when the list is filtered) |
| `Backspace` | Erase typed digits |
| `A` | Toggle the **hide already-migrated** filter |
| `R` | Refresh status from the server |
| `Q` / `Esc` | Cancel |
| `Enter` | Confirm the highlighted project and start the mirror |

### Source / target selection

- If the config contains a collection named **`MDR-GAMS-ADO`**, it is auto-selected as the target.
- If exactly one other collection is configured, it is auto-selected as the source.
- Otherwise you get the standard interactive collection picker for whichever side is missing.

### Run it

```powershell
# Interactive — the normal way (also reachable via Start-Menu option [13])
./Invoke-ProjectMigration.ps1 -ConfigPath ./config/migration-config.json

# Dry run — show the picker, walk the mirror flow, but skip clones/pushes
./Invoke-ProjectMigration.ps1 -ConfigPath ./config/migration-config.json -DryRun

# Fully scripted — pick the project up front, skip the final confirm prompt
./Invoke-ProjectMigration.ps1 -ConfigPath ./config/migration-config.json `
    -SourceCollection GAMS-GIT-Repos -TargetCollection MDR-GAMS-ADO -Force
```

Under the hood this hands off to [`Mirror-AdoCollection.ps1`](Mirror-AdoCollection.ps1) with `-SourceProject <name>`, so you get the same nested progress bars, rollup summary table, and manifest CSV as a full-collection mirror — just scoped to one project.

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
| `Mirror-AdoCollection.ps1` | Mirrors every project & Git repo from one collection/org to another (cross-server supported) | `-Interactive` |
| `Invoke-ProjectMigration.ps1` | Interactive single-project picker that shows migration status and hands off to the mirror | Yes (always) |
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

1. **Setup wizard** — run `./New-MigrationConfig.ps1` or pick option **[1]** from the main menu. It walks you through each setting and tests your ADO connections. For every collection you add, the wizard asks **“Is this an Azure DevOps Services org (hosted at `https://dev.azure.com/<name>`)?”** — answer `y` to store a `serverUrl` override (e.g. for `MDR-GAMS-ADO`), or accept the default `n` to keep using the on-prem `adoServerUrl`. Existing Services-org entries default to **Y** when you re-run the wizard so updates don’t silently demote them back to on-prem.
2. **Manual** — copy [`config/migration-config.example.json`](config/migration-config.example.json) to `config/migration-config.json` and edit it.

Key settings:

| Setting | Description |
|---|---|
| `adoServerUrl` | Default ADO server base URL (e.g. `https://ado.mcdermott.com`). Each collection inherits this unless it sets its own `serverUrl`. |
| `collections` | Map of collection (or ADO Services org) names. Each entry takes a `pat`, an optional `description`, and an optional `serverUrl` override (set this to `https://dev.azure.com` for an ADO Services org). |
| `outputDirectory` | Where converted Git repos are written (default: `./output`) |
| `logDirectory` | Where log files are written (default: `./logs`) |
| `gitTfsPath` | Path to git-tfs if it's not in your PATH |
| `authorMappingFile` | Optional CSV mapping TFVC usernames → Git authors |
| `github.enterpriseUrl` | GitHub Enterprise URL (e.g. `https://github.mcdermott.com`) |
| `github.pat` | GitHub PAT with `repo` scope |
| `github.defaultOrg` | Default GitHub organization for new repos |
| `migrationDefaults.timeoutMinutes` | Optional hard timeout per conversion in minutes (example: `120`) |
| `migrationDefaults.stallTimeoutMinutes` | Optional no-output stall timeout in minutes (example: `30`) |

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
├── Mirror-AdoCollection.ps1          # Full-collection Git mirror (cross-server)
├── Invoke-ProjectMigration.ps1       # Interactive single-project picker (option [13])
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

All operations write timestamped logs to `./logs/`. View recent logs from the main menu (option **[11]**) or browse the directory directly. Each script also supports `-Verbose` for detailed console output. Git push command logging now redacts PAT credentials in authenticated URLs.

## Troubleshooting

Common issues and what to do:

| Symptom | Likely Cause | Fix |
|---|---|---|
| "Authentication failed — your PAT may be expired" | PAT expired or wrong | Generate a new PAT in ADO and update your config |
| "Access denied — your PAT doesn't have the required permissions" | Missing PAT scopes | Ensure `Code (Read & Write)` and `Project and Team (Read)` |
| "Basic authentication requires a secure connection" | Server URL uses `http://` | Change `adoServerUrl` in your config to `https://` |
| TF400959 "limit of 248 characters" / path too long errors | Windows MAX_PATH limit hit by deeply nested TFVC files | The toolkit now automatically uses `subst` to create a short drive-letter path during cloning, which dramatically reduces path length. This happens transparently — no manual steps needed. If you still hit issues: **1)** Set `outputDirectory` to something very short like `C:\M`; **2)** Enable long paths system-wide: `reg add HKLM\SYSTEM\CurrentControlSet\Control\FileSystem /v LongPathsEnabled /t REG_DWORD /d 1 /f` (admin + reboot); **3)** Ensure `core.longpaths` is set (the toolkit does this automatically). If paths in TFVC exceed ~200 characters even after the drive root, consider using the `Split-TfvcToGitRepos.ps1` script to clone only the subfolder you need rather than the entire project tree. |
| "Cannot reach the ADO server" | Network/VPN issue | Check VPN connection and server URL in config |
| "git-tfs not found" | Tool not installed | Run `Install-Prerequisites.ps1` for instructions |
| "Missing Required Module: ImportExcel" | ImportExcel not installed | Run `Install-Module ImportExcel -Scope CurrentUser -Force` |
| Conversion seems frozen | Large repo, blocked prompt, or no git-tfs output | Batch mode now auto-cleans output folders (no prompt), detects stalls/timeouts, marks item `TimedOut`, and continues. Tune `timeoutMinutes` / `stallTimeoutMinutes` if needed. |

### Resuming After Failures (Skip & Continue)

When the Excel-driven migration (`Invoke-ExcelMigration.ps1`) encounters a failure — especially a **PathTooLong** error — it records the status in the output manifest CSV. You can resume from where you left off:

1. **Run the migration** — it produces a manifest CSV (e.g. `output/excel-migration-manifest-20260514-093000.csv`)
2. **Open the manifest CSV** in Excel and review items with Status = `Failed` or `PathTooLong`
3. **Change the Status** of any items you want to skip permanently to `Skipped`
4. **Re-run with `-ResumeManifest`** — items marked `Success`, `Skipped`, or `PathTooLong` are automatically excluded. Items marked `TimedOut` are retried by default so they get another chance:

```powershell
./Invoke-ExcelMigration.ps1 -ConfigPath ./config/migration-config.json -Interactive `
    -ResumeManifest ./output/excel-migration-manifest-20260514-093000.csv
```

### Per-Folder Targeted Cloning

When the Excel spreadsheet marks individual folders for migration (Repo or Folder = "Folder"), the toolkit clones **each folder independently** from its specific TFVC path (e.g. `$/InformationTechnology/CMeR`) rather than cloning the entire parent repository. This means:

- A deeply nested folder like `PICASharepointSolution` won't prevent `CMeR`, `Coreworx`, or other sibling folders from migrating
- Each folder is isolated — if one fails, the rest continue
- Cloning is faster because only the relevant folder's files and history are downloaded
