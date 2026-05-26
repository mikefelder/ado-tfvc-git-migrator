#Requires -Version 7.0
<#
.SYNOPSIS
    Processes the MDR-4ADO-AllProjects Excel file to batch migrate/split/skip TFVC repos
    and mirror existing Git repos in a single second-pass run.

.DESCRIPTION
    Reads the "GAMS-Repos-App-Folder level" worksheet from MDR-4ADO-AllProjects.xlsx
    and processes each row based on the Recommendation, Repos Type and Repo/Folder columns:

    - Recommendation = "Archive*"                                       → Skip (no migration)
    - Repo or Folder = "Repo"   + Recommendation = "Migrate*" + TFVC    → Convert entire repo from TFVC to Git, move to -TargetCollection / -TargetProject
    - Repo or Folder = "Folder" + Recommendation = "Migrate*" + TFVC    → Spin out folder from parent repo into new Git repo, move to -TargetCollection / -TargetProject
    - Repo or Folder = "Repo"   + Recommendation = "Migrate*" + Git     → Mirror existing Git repo (clone --mirror + push --mirror) into -GitTargetCollection (default 'GAMS-GIT-Repos'); target project defaults to the source project name unless -GitTargetProject is given
    - Repo or Folder = "Folder" + Recommendation = "Migrate*" + Git     → Skipped with a clear reason (sub-tree extraction from a Git repo is out of scope here)

    Each folder is cloned independently from its own TFVC path rather than cloning
    the entire parent repo. This avoids path-too-long failures caused by deeply
    nested files in sibling folders.

    Produces a comprehensive manifest of all actions taken.

.PARAMETER ConfigPath
    Path to migration-config.json.

.PARAMETER ExcelPath
    Path to the MDR-4ADO-AllProjects.xlsx file.

.PARAMETER WorksheetName
    Worksheet to read. Default: "GAMS-Repos-App-Folder level".

.PARAMETER TargetCollection
    Target ADO collection for **TFVC** migrations. Required only if the spreadsheet
    contains any TFVC rows marked for migration.

.PARAMETER TargetProject
    Target ADO project for **TFVC** migrations. Required only if the spreadsheet
    contains any TFVC rows marked for migration.

.PARAMETER GitTargetCollection
    Target ADO collection / Services org for **Git** mirror rows. Defaults to
    'GAMS-GIT-Repos'. Must exist in the config. Only used when the spreadsheet
    has at least one Git source row.

.PARAMETER GitTargetProject
    Optional target project for **Git** mirror rows. When empty (default), the
    source project name is preserved and the target project is auto-created if
    it does not exist.

.PARAMETER ResumeManifest
    Path to a previous migration manifest CSV. Items already marked
    'Success', 'Skipped', or 'PathTooLong' in the manifest will be
    skipped automatically. Edit the CSV to change a failed item's
    Status to 'Skipped' to exclude it from future runs.

.PARAMETER Interactive
    Launch interactive mode — step-by-step guided flow with preview and confirmation.

.EXAMPLE
    # Interactive mode (recommended)
    ./Invoke-ExcelMigration.ps1 -ConfigPath ./config/migration-config.json -Interactive

.EXAMPLE
    # Resume from a previous run — skips items already completed or marked Skipped
    ./Invoke-ExcelMigration.ps1 -ConfigPath ./config/migration-config.json -Interactive `
        -ResumeManifest ./output/excel-migration-manifest-20260514-093000.csv

.EXAMPLE
    # Direct — TFVC rows only
    ./Invoke-ExcelMigration.ps1 -ConfigPath ./config/migration-config.json `
        -ExcelPath ./excel-docs/MDR-4ADO-AllProjects.xlsx `
        -TargetCollection "ModernApps" -TargetProject "Platform"

.EXAMPLE
    # Direct — second-pass run for Git-source rows, mirrored into GAMS-GIT-Repos
    ./Invoke-ExcelMigration.ps1 -ConfigPath ./config/migration-config.json `
        -ExcelPath ./excel-docs/MDR-4ADO-AllProjects.xlsx `
        -GitTargetCollection "GAMS-GIT-Repos"
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory)]
    [string]$ConfigPath,

    [string]$ExcelPath,

    [string]$WorksheetName = 'GAMS-Repos-App-Folder level',

    [string]$TargetCollection,

    [string]$TargetProject,

    [string]$GitTargetCollection = 'GAMS-GIT-Repos',

    [string]$GitTargetProject,

    [string]$ResumeManifest,

    [switch]$Interactive
)

$ErrorActionPreference = 'Stop'
Import-Module "$PSScriptRoot/modules/AdoTfvcMigrator.psm1" -Force

function Resolve-TfvcFolderPath {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$SourceProject,
        [Parameter(Mandatory)]
        [string]$SourceRepo,
        [Parameter(Mandatory)]
        [string]$FolderName
    )

    $normalized = ($FolderName -replace '\\', '/').Trim()
    $normalized = $normalized.Trim('/')

    if ($normalized -match '^\$/') {
        return $normalized
    }

    if ($normalized -like "$SourceProject/*") {
        return "`$/$normalized"
    }

    if ($normalized -like "$SourceRepo/*") {
        return "`$/$SourceProject/$normalized"
    }

    return "`$/$SourceProject/$SourceRepo/$normalized"
}

function Test-TfvcPathExists {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$ServerUrl,
        [Parameter(Mandatory)]
        [string]$Collection,
        [Parameter(Mandatory)]
        [string]$Pat,
        [Parameter(Mandatory)]
        [string]$TfvcPath
    )

    try {
        $items = @(Get-TfvcItems -ServerUrl $ServerUrl -Collection $Collection -Pat $Pat -ScopePath $TfvcPath -RecursionLevel 1)
        return [bool]($items | Where-Object { $_.path -eq $TfvcPath } | Select-Object -First 1)
    }
    catch {
        return $false
    }
}

# ─── Check for ImportExcel Module ──────────────────────────────────────────────

if (-not (Get-Module -ListAvailable -Name ImportExcel)) {
    Write-Host ""
    Write-Host "  ┌──────────────────────────────────────────────────────────┐" -ForegroundColor Red
    Write-Host "  │  Missing Required Module: ImportExcel                    │" -ForegroundColor Red
    Write-Host "  │                                                          │" -ForegroundColor Red
    Write-Host "  │  This module is needed to read the Excel spreadsheet.    │" -ForegroundColor Red
    Write-Host "  │                                                          │" -ForegroundColor Red
    Write-Host "  │  To install it, run this command:                        │" -ForegroundColor Red
    Write-Host "  │  Install-Module ImportExcel -Scope CurrentUser -Force    │" -ForegroundColor Yellow
    Write-Host "  │                                                          │" -ForegroundColor Red
    Write-Host "  │  Then come back and try again.                           │" -ForegroundColor Red
    Write-Host "  └──────────────────────────────────────────────────────────┘" -ForegroundColor Red
    Write-Host ""
    return
}
Import-Module ImportExcel

# ─── Load Config ───────────────────────────────────────────────────────────────

$config = Read-MigrationConfig -ConfigPath $ConfigPath
$logFile = Initialize-MigrationLog -LogDirectory $config.logDirectory -ScriptName 'ExcelMigration'

# ─── Load Resume Manifest (if provided) ───────────────────────────────────────

$resumeSkipSet = @{}
if ($ResumeManifest) {
    if (-not (Test-Path $ResumeManifest)) {
        Write-Host "  Resume manifest not found: $ResumeManifest" -ForegroundColor Red
        return
    }

    $previousManifest = Import-Csv $ResumeManifest
    $resumeStatuses = @('Success', 'Skipped', 'PathTooLong')

    foreach ($prev in $previousManifest) {
        if ($prev.Status -in $resumeStatuses) {
            # Key by Collection|Project|Repo|Folder to uniquely identify each item
            $key = "$($prev.Collection)|$($prev.Project)|$($prev.Repo)|$($prev.Folder)"
            $resumeSkipSet[$key] = $prev.Status
        }
    }

    $resumeCount = $resumeSkipSet.Count
    Write-MigrationLog -Message "Resume manifest loaded: $resumeCount item(s) will be skipped (Success/Skipped/PathTooLong)" -LogFile $logFile -Level INFO

    if ($Interactive) {
        Write-Host "  ┌─ Resuming from previous run ──────────────────────────┐" -ForegroundColor Green
        Write-Host "  │  Loaded: $ResumeManifest" -ForegroundColor Green
        Write-Host "  │  $resumeCount item(s) already completed or skipped    " -ForegroundColor Green
        Write-Host "  │  will be excluded from this run.                       " -ForegroundColor Green
        Write-Host "  └───────────────────────────────────────────────────────┘" -ForegroundColor Green
        Write-Host ""
        Write-Host "  Tip: To skip an item that failed, open the manifest CSV" -ForegroundColor DarkGray
        Write-Host "  and change its Status from 'Failed' or 'PathTooLong' to" -ForegroundColor DarkGray
        Write-Host "  'Skipped', then re-run with the same -ResumeManifest."   -ForegroundColor DarkGray
        Write-Host ""
    }
}

# ─── Step 1: Locate the Excel File ────────────────────────────────────────────

if ($Interactive) {
    Show-MenuHeader -Title "Excel-Driven Batch Migration"

    Write-Host "  This tool will:" -ForegroundColor White
    Write-Host ""
    Write-Host "    1. Read the MDR-4ADO-AllProjects spreadsheet" -ForegroundColor DarkGray
    Write-Host "    2. Show you exactly what will happen to each repository" -ForegroundColor DarkGray
    Write-Host "    3. Ask for your confirmation before making any changes" -ForegroundColor DarkGray
    Write-Host "    4. Migrate the repositories and produce a full report" -ForegroundColor DarkGray
    Write-Host ""
    Write-Host "  No changes are made until you explicitly confirm." -ForegroundColor Green
    Write-Host ""
    Write-Host "  ─────────────────────────────────────────────────────" -ForegroundColor DarkGray
    Write-Host "  Step 1 of 4: Locate the spreadsheet" -ForegroundColor Cyan
    Write-Host "  ─────────────────────────────────────────────────────" -ForegroundColor DarkGray
    Write-Host ""
}

if (-not $ExcelPath) {
    $defaultExcel = Join-Path $PSScriptRoot 'excel-docs/MDR-4ADO-AllProjects.xlsx'
    if (Test-Path $defaultExcel) {
        if ($Interactive) {
            Write-Host "  ✓ Found the spreadsheet automatically:" -ForegroundColor Green
            Write-Host "    $defaultExcel" -ForegroundColor White
            Write-Host ""
        }
        $ExcelPath = $defaultExcel
    }
    elseif ($Interactive) {
        Write-Host "  The spreadsheet was not found in the expected location." -ForegroundColor Yellow
        Write-Host "  Please enter the full path to MDR-4ADO-AllProjects.xlsx:" -ForegroundColor White
        Write-Host ""
        Write-Host "  Path: " -ForegroundColor Yellow -NoNewline
        $ExcelPath = Read-Host
    }
}

if (-not $ExcelPath -or -not (Test-Path $ExcelPath)) {
    Write-Host ""
    Write-Host "  Could not find the spreadsheet file." -ForegroundColor Red
    Write-Host "  Make sure the file 'MDR-4ADO-AllProjects.xlsx' is in the excel-docs folder." -ForegroundColor Yellow
    Write-Host ""
    return
}

# ─── Step 2: Read & Analyze the Spreadsheet ──────────────────────────────────

if ($Interactive) {
    Write-Host "  ─────────────────────────────────────────────────────" -ForegroundColor DarkGray
    Write-Host "  Step 2 of 4: Reading the spreadsheet..." -ForegroundColor Cyan
    Write-Host "  ─────────────────────────────────────────────────────" -ForegroundColor DarkGray
    Write-Host ""
}

if ($Interactive) {
    Write-Host ""
    Write-Host "  ─────────────────────────────────────────────────────" -ForegroundColor DarkGray
    Write-Host "  Step 3 of 4: Reading the spreadsheet..." -ForegroundColor Cyan
    Write-Host "  ─────────────────────────────────────────────────────" -ForegroundColor DarkGray
    Write-Host ""
}

Write-MigrationLog -Message "Reading Excel file: $ExcelPath (sheet: $WorksheetName)" -LogFile $logFile -Level INFO

try {
    $excelData = Import-Excel -Path $ExcelPath -WorksheetName $WorksheetName
}
catch {
    Write-Host "  Could not read the spreadsheet." -ForegroundColor Red
    Write-Host "  Error: $($_.Exception.Message)" -ForegroundColor Red
    Write-Host ""
    Write-Host "  Make sure the file is a valid Excel file and is not open in another program." -ForegroundColor Yellow
    Write-Host ""
    return
}

if (-not $excelData -or $excelData.Count -eq 0) {
    Write-Host "  The spreadsheet appears to be empty." -ForegroundColor Red
    Write-Host "  Expected data in worksheet: '$WorksheetName'" -ForegroundColor Yellow
    Write-Host ""
    return
}

Write-MigrationLog -Message "Loaded $($excelData.Count) rows from Excel" -LogFile $logFile -Level INFO

# ─── Classify Every Row ───────────────────────────────────────────────────────

$migrateRows = [System.Collections.ArrayList]::new()
$archiveRows = [System.Collections.ArrayList]::new()
$skipRows = [System.Collections.ArrayList]::new()
$dataWarnings = [System.Collections.ArrayList]::new()

$excelRowIndex = 1  # Header is row 1; first data row starts at 2
$recommendationCol = 8  # Column H = Recommendation

foreach ($row in $excelData) {
    $excelRowIndex++
    $recommendation = ($row.Recommendation ?? '').ToString().Trim()
    $repoOrFolder = ($row.'Repo or Floder' ?? '').ToString().Trim()
    $reposType = ($row.'Repos Type' ?? '').ToString().Trim()
    $collection = ($row.Collection ?? '').ToString().Trim()
    $project = ($row.'Projects ' ?? $row.Projects ?? '').ToString().Trim()
    $repoName = ($row.Repos ?? '').ToString().Trim()
    $folderName = ($row.'Applications/Folders' ?? '').ToString().Trim()

    # Build a standard entry
    $entry = [PSCustomObject]@{
        RowNo          = $row.No
        ExcelRow       = $excelRowIndex
        Collection     = $collection
        Project        = $project
        Repo           = $repoName
        Folder         = $folderName
        RepoType       = $reposType
        RepoOrFolder   = $repoOrFolder
        Recommendation = $recommendation
        Action         = ''
        NewRepoName    = ''
        Reason         = ''
        Status         = ''
        Error          = ''
    }

    # Skip rows with no collection/project data
    if (-not $collection -or -not $project) {
        $entry.Action = 'Skip'
        $entry.Reason = 'Missing collection or project name in spreadsheet'
        $entry.Status = 'Skipped'
        [void]$skipRows.Add($entry)
        continue
    }

    # Exclude BuildProcessTemplates (matches bare name or any TFVC path ending in /BuildProcessTemplates)
    if ($folderName -eq 'BuildProcessTemplates' -or $folderName -like '*/BuildProcessTemplates') {
        $entry.Action = 'Skip'
        $entry.Reason = 'BuildProcessTemplates (system folder, excluded automatically)'
        $entry.Status = 'Skipped'
        [void]$skipRows.Add($entry)
        continue
    }

    # Skip rows already marked as done/completed from a previous migration run
    if ($recommendation -match '(?i)^(done|completed)$') {
        $entry.Action = 'Skip'
        $entry.Reason = "Already completed (marked '$recommendation' in spreadsheet)"
        $entry.Status = 'Skipped'
        [void]$skipRows.Add($entry)
        continue
    }

    # Silently ignore rows whose collection is not in the config (no PAT = can't migrate)
    if (-not $config.collections[$collection]) {
        continue
    }

    # ── Resume: skip items already completed or manually marked Skipped ──
    if ($resumeSkipSet.Count -gt 0) {
        $resumeKey = "$collection|$project|$repoName|$folderName"
        if ($resumeSkipSet.ContainsKey($resumeKey)) {
            $prevStatus = $resumeSkipSet[$resumeKey]
            $entry.Action = 'Skip'
            $entry.Reason = "Previously $prevStatus (from resume manifest)"
            $entry.Status = 'Skipped'
            [void]$skipRows.Add($entry)
            continue
        }
    }

    # Classify by recommendation
    if ($recommendation -match '(?i)^archive') {
        $entry.Action = 'Archive'
        $entry.Reason = "Marked as '$recommendation' — will be skipped"
        $entry.Status = 'Skipped'
        [void]$archiveRows.Add($entry)
    }
    elseif ($recommendation -match '(?i)^migrate') {
        $isGit = $reposType -match '(?i)^git'

        if ($repoOrFolder -eq 'Folder' -and $isGit) {
            # Sub-folder extraction from an existing Git repo is out of scope here —
            # clone the repo locally and use git filter-repo / git subtree split instead.
            $entry.Action = 'Skip'
            $entry.Reason = "Git sub-folder extraction not supported by this script — clone '$repoName' and split '$folderName' manually"
            $entry.Status = 'Skipped'
            [void]$skipRows.Add($entry)
        }
        elseif ($repoOrFolder -eq 'Folder') {
            $entry.Action = 'SpinOutFolder'
            $entry.NewRepoName = "$($repoName)_$($folderName)" -replace '[^a-zA-Z0-9_-]', '-'
            $entry.Reason = "Extract folder '$folderName' from repo '$repoName' → new repo '$($entry.NewRepoName)'"
            $entry.Status = 'Pending'
            [void]$migrateRows.Add($entry)
        }
        else {
            # "Repo" means migrate the entire repo as one unit.
            # Multiple spreadsheet rows may reference the same repo (one per subfolder) —
            # only add it once; skip duplicate rows for the same Collection/Project/Repo.
            $repoKey = "$($collection)|$($project)|$($repoName)"
            $alreadyAdded = $migrateRows | Where-Object {
                $_.Action -in @('MigrateRepo', 'MirrorGitRepo') -and
                "$($_.Collection)|$($_.Project)|$($_.Repo)" -eq $repoKey
            }
            if ($alreadyAdded) {
                $entry.Action = 'Skip'
                $entry.Reason = "Duplicate row — repo '$repoName' is already queued for migration"
                $entry.Status = 'Skipped'
                [void]$skipRows.Add($entry)
            }
            elseif ($isGit) {
                # Existing Git repo → mirror it (clone --mirror + push --mirror) into
                # the Git target collection. Target project defaults to the source
                # project name and is auto-created if it doesn't already exist.
                $entry.Action = 'MirrorGitRepo'
                $entry.NewRepoName = $repoName -replace '[^a-zA-Z0-9_-]', '-'
                $entry.Reason = "Mirror existing Git repo '$repoName' from $collection/$project → $GitTargetCollection"
                $entry.Status = 'Pending'
                [void]$migrateRows.Add($entry)
            }
            else {
                $entry.Action = 'MigrateRepo'
                $entry.NewRepoName = $repoName -replace '[^a-zA-Z0-9_-]', '-'
                $entry.Reason = "Migrate entire repo '$repoName' to Git"
                $entry.Status = 'Pending'
                [void]$migrateRows.Add($entry)
            }
        }
    }
    else {
        $entry.Action = 'Skip'
        $entry.Reason = if ($recommendation) { "Recommendation '$recommendation' not recognized" } else { 'No recommendation specified' }
        $entry.Status = 'Skipped'
        [void]$skipRows.Add($entry)
    }
}
# ─── Resolve Duplicate Repo Names ─────────────────────────────────────────

$nameCount = @{}
foreach ($item in $migrateRows) {
    $name = $item.NewRepoName
    if (-not $nameCount.ContainsKey($name)) {
        $nameCount[$name] = [System.Collections.ArrayList]::new()
    }
    [void]$nameCount[$name].Add($item)
}

foreach ($kvp in $nameCount.GetEnumerator()) {
    if ($kvp.Value.Count -gt 1) {
        # First occurrence keeps the base name, subsequent ones get _1, _2, etc.
        for ($i = 1; $i -lt $kvp.Value.Count; $i++) {
            $kvp.Value[$i].NewRepoName = "$($kvp.Key)_$i"
            $kvp.Value[$i].Reason = $kvp.Value[$i].Reason -replace [regex]::Escape($kvp.Key), $kvp.Value[$i].NewRepoName
        }
    }
}

# ─── Step 3: Choose Target(s) — conditional on what was classified ───────────
#
# Only prompt / require targets for the kinds of work the spreadsheet actually
# has. A pure TFVC pass doesn't need a Git target; a pure Git mirror pass
# doesn't need a TFVC target.

$tfvcWorkItems = @($migrateRows | Where-Object { $_.Action -in @('MigrateRepo', 'SpinOutFolder') })
$gitMirrorItems = @($migrateRows | Where-Object { $_.Action -eq 'MirrorGitRepo' })

if ($Interactive -and ($tfvcWorkItems.Count -gt 0 -or $gitMirrorItems.Count -gt 0)) {
    Write-Host ""
    Write-Host "  ─────────────────────────────────────────────────────" -ForegroundColor DarkGray
    Write-Host "  Step 3 of 4: Where should migrated repos go?" -ForegroundColor Cyan
    Write-Host "  ─────────────────────────────────────────────────────" -ForegroundColor DarkGray
    Write-Host ""
}

# TFVC target — only if there are TFVC migrations or folder spin-outs queued
if ($tfvcWorkItems.Count -gt 0) {
    if ($Interactive) {
        Write-Host "  TFVC conversions ($($tfvcWorkItems.Count) item(s)) need a target collection + project." -ForegroundColor White
        Write-Host ""
    }

    if (-not $TargetCollection -and $Interactive) {
        $TargetCollection = Select-AdoCollection -Config $config -Prompt 'Select the target collection for TFVC migrations'
        if (-not $TargetCollection) {
            Write-Host "  Cancelled." -ForegroundColor Yellow
            return
        }
    }

    if (-not $TargetProject -and $Interactive) {
        Write-Host ""
        Write-Host "  Enter the target project name for TFVC migrations: " -ForegroundColor Yellow -NoNewline
        $TargetProject = Read-Host
        if (-not $TargetProject) {
            Write-Host "  Cancelled." -ForegroundColor Yellow
            return
        }
    }

    if (-not $TargetCollection -or -not $TargetProject) {
        Write-Host "  -TargetCollection and -TargetProject are required when the spreadsheet contains TFVC migration rows." -ForegroundColor Red
        return
    }
}

# Git target — only if there are Git mirror rows queued
if ($gitMirrorItems.Count -gt 0) {
    if ($Interactive) {
        Write-Host ""
        Write-Host "  Existing Git repos ($($gitMirrorItems.Count) item(s)) will be mirrored into a target collection." -ForegroundColor White
        Write-Host "  Source project names are preserved by default and auto-created if missing." -ForegroundColor DarkGray
        Write-Host ""

        $gitDefault = if ($GitTargetCollection) { $GitTargetCollection } else { 'GAMS-GIT-Repos' }
        Write-Host "  Git target collection [default: $gitDefault]: " -ForegroundColor Yellow -NoNewline
        $entered = Read-Host
        if ($entered) { $GitTargetCollection = $entered.Trim() }
        else { $GitTargetCollection = $gitDefault }

        if (-not $GitTargetProject) {
            Write-Host "  Git target project (blank = preserve each source project name): " -ForegroundColor Yellow -NoNewline
            $enteredProj = Read-Host
            if ($enteredProj) { $GitTargetProject = $enteredProj.Trim() }
        }
    }

    if (-not $GitTargetCollection) {
        Write-Host "  -GitTargetCollection is required (default 'GAMS-GIT-Repos') when the spreadsheet contains Git mirror rows." -ForegroundColor Red
        return
    }

    if (-not $config.collections.ContainsKey($GitTargetCollection)) {
        Write-Host ""
        Write-Host "  ✗ Git target collection '$GitTargetCollection' is not in your configuration." -ForegroundColor Red
        Write-Host "    Add it with the Setup Wizard (option 1 on the main menu) and ensure its PAT" -ForegroundColor Yellow
        Write-Host "    has Code (Read & Write) AND Project and Team (Read, Write & Manage) scopes." -ForegroundColor Yellow
        Write-Host ""
        return
    }
}

# ─── Show the Preview ─────────────────────────────────────────────────────────

$repoMigrations = @($migrateRows | Where-Object { $_.Action -eq 'MigrateRepo' })
$folderSpinouts = @($migrateRows | Where-Object { $_.Action -eq 'SpinOutFolder' })
$gitMirrors = @($migrateRows | Where-Object { $_.Action -eq 'MirrorGitRepo' })
$tfvcConversions = @($migrateRows | Where-Object { $_.RepoType -eq 'TFVC' })

Write-Host ""
Write-Host "  ╔═══════════════════════════════════════════════════════╗" -ForegroundColor Cyan
Write-Host "  ║              Migration Preview                        ║" -ForegroundColor Cyan
Write-Host "  ╚═══════════════════════════════════════════════════════╝" -ForegroundColor Cyan
Write-Host ""
Write-Host "  The spreadsheet has $($excelData.Count) rows. Here is what will happen:" -ForegroundColor White
Write-Host ""
Write-Host "    ● $($migrateRows.Count) repositories will be migrated" -ForegroundColor Green
if ($repoMigrations.Count -gt 0) {
    Write-Host "        ├─ $($repoMigrations.Count) full repo migration(s)" -ForegroundColor DarkGray
}
if ($folderSpinouts.Count -gt 0) {
    Write-Host "        ├─ $($folderSpinouts.Count) folder(s) will be extracted into standalone repos" -ForegroundColor DarkGray
}
if ($gitMirrors.Count -gt 0) {
    Write-Host "        └─ $($gitMirrors.Count) existing Git repo(s) will be mirrored (clone --mirror / push --mirror)" -ForegroundColor DarkGray
}
if ($tfvcConversions.Count -gt 0) {
    Write-Host "        ($($tfvcConversions.Count) of these need to be converted from TFVC to Git format)" -ForegroundColor DarkGray
}
Write-Host "    ● $($archiveRows.Count) repositories are marked for archive (will be skipped)" -ForegroundColor Yellow
Write-Host "    ● $($skipRows.Count) rows will be skipped (system folders, missing data, etc.)" -ForegroundColor DarkGray
Write-Host ""
if ($tfvcWorkItems.Count -gt 0) {
    Write-Host "    TFVC destination: $TargetCollection / $TargetProject" -ForegroundColor White
}
if ($gitMirrors.Count -gt 0) {
    $gitProjLabel = if ($GitTargetProject) { $GitTargetProject } else { '<preserve source project>' }
    Write-Host "    Git mirror destination: $GitTargetCollection / $gitProjLabel" -ForegroundColor White
}
Write-Host ""

# Show data warnings
if ($dataWarnings.Count -gt 0) {
    Write-Host "  ┌─ Warnings Found ──────────────────────────────────────┐" -ForegroundColor Yellow
    foreach ($w in $dataWarnings | Select-Object -First 10) {
        Write-Host "  │  ⚠ $w" -ForegroundColor Yellow
    }
    if ($dataWarnings.Count -gt 10) {
        Write-Host "  │  ... and $($dataWarnings.Count - 10) more warnings" -ForegroundColor Yellow
    }
    Write-Host "  └───────────────────────────────────────────────────────┘" -ForegroundColor Yellow
    Write-Host ""
}

# Show detailed preview by collection
$migrateByCollection = $migrateRows | Group-Object -Property Collection

Write-Host "  ── What will be migrated ─────────────────────────────" -ForegroundColor DarkGray
Write-Host ""

foreach ($grp in $migrateByCollection) {
    Write-Host "  Collection: $($grp.Name)" -ForegroundColor White
    foreach ($item in $grp.Group) {
        if ($item.Action -eq 'MigrateRepo') {
            $typeNote = if ($item.RepoType -eq 'TFVC') { ' [convert to Git]' } else { '' }
            Write-Host "    → Migrate repo: $($item.Project)/$($item.Repo)$typeNote" -ForegroundColor DarkGray
        }
        elseif ($item.Action -eq 'MirrorGitRepo') {
            $tgtProj = if ($GitTargetProject) { $GitTargetProject } else { $item.Project }
            Write-Host "    → Mirror Git repo: $($item.Project)/$($item.Repo) → $GitTargetCollection/$tgtProj/$($item.NewRepoName)" -ForegroundColor DarkGray
        }
        else {
            Write-Host "    → Extract folder: $($item.Project)/$($item.Repo)/$($item.Folder) → new repo '$($item.NewRepoName)'" -ForegroundColor DarkGray
        }
    }
    Write-Host ""
}

if ($archiveRows.Count -gt 0) {
    Write-Host "  ── What will be skipped (archive) ────────────────────" -ForegroundColor DarkGray
    Write-Host ""
    $archiveByCollection = $archiveRows | Group-Object -Property Collection
    foreach ($grp in ($archiveByCollection | Select-Object -First 5)) {
        Write-Host "  Collection: $($grp.Name) ($($grp.Count) repos)" -ForegroundColor DarkGray
    }
    if ($archiveByCollection.Count -gt 5) {
        Write-Host "  ... and $($archiveByCollection.Count - 5) more collections" -ForegroundColor DarkGray
    }
    Write-Host ""
}

# Always save the preview manifest
$allPreviewActions = [System.Collections.ArrayList]::new()
foreach ($r in $migrateRows) { [void]$allPreviewActions.Add($r) }
foreach ($r in $archiveRows) { [void]$allPreviewActions.Add($r) }
foreach ($r in $skipRows) { [void]$allPreviewActions.Add($r) }

$previewManifestPath = Join-Path $config.outputDirectory "excel-migration-PREVIEW-$(Get-Date -Format 'yyyyMMdd-HHmmss').csv"
$allPreviewActions | Export-Csv -Path $previewManifestPath -NoTypeInformation -Encoding utf8
Write-MigrationLog -Message "Preview manifest written to: $previewManifestPath" -LogFile $logFile -Level SUCCESS

Write-Host "  A detailed preview has been saved to:" -ForegroundColor DarkGray
Write-Host "  $previewManifestPath" -ForegroundColor White
Write-Host "  (You can open this CSV in Excel to review all $($excelData.Count) rows)" -ForegroundColor DarkGray
Write-Host ""

# ─── Step 4: Confirm and Execute ──────────────────────────────────────────────

if ($Interactive) {
    Write-Host "  ─────────────────────────────────────────────────────" -ForegroundColor DarkGray
    Write-Host "  Step 4 of 4: Confirm and run" -ForegroundColor Cyan
    Write-Host "  ─────────────────────────────────────────────────────" -ForegroundColor DarkGray
    Write-Host ""

    if ($migrateRows.Count -eq 0) {
        Write-Host "  There are no repositories to migrate." -ForegroundColor Yellow
        Write-Host "  All rows are either marked for archive or were skipped." -ForegroundColor DarkGray
        Write-Host ""
        return
    }

    Write-Host "  ┌──────────────────────────────────────────────────────────┐" -ForegroundColor Yellow
    Write-Host "  │                                                          │" -ForegroundColor Yellow
    Write-Host "  │  Ready to migrate $($migrateRows.Count.ToString().PadRight(4)) repositories.                    │" -ForegroundColor Yellow
    Write-Host "  │                                                          │" -ForegroundColor Yellow
    Write-Host "  │  This will:                                              │" -ForegroundColor Yellow
    if ($tfvcConversions.Count -gt 0) {
        Write-Host "  │    • Convert $($tfvcConversions.Count.ToString().PadRight(4)) repos from TFVC to Git format     │" -ForegroundColor Yellow
    }
    if ($folderSpinouts.Count -gt 0) {
        Write-Host "  │    • Extract $($folderSpinouts.Count.ToString().PadRight(4)) folders into standalone repos      │" -ForegroundColor Yellow
    }
    if ($gitMirrors.Count -gt 0) {
        Write-Host "  │    • Mirror $($gitMirrors.Count.ToString().PadRight(4)) existing Git repos                   │" -ForegroundColor Yellow
    }
    if ($tfvcWorkItems.Count -gt 0) {
        Write-Host "  │    • TFVC dest: $($TargetCollection)/$($TargetProject)".PadRight(60) + "│" -ForegroundColor Yellow
    }
    if ($gitMirrors.Count -gt 0) {
        $gitProjLabel2 = if ($GitTargetProject) { $GitTargetProject } else { '<keep source>' }
        Write-Host "  │    • Git dest:  $($GitTargetCollection)/$gitProjLabel2".PadRight(60) + "│" -ForegroundColor Yellow
    }
    Write-Host "  │                                                          │" -ForegroundColor Yellow
    Write-Host "  │  This may take a while for large repositories.           │" -ForegroundColor Yellow
    Write-Host "  │                                                          │" -ForegroundColor Yellow
    Write-Host "  └──────────────────────────────────────────────────────────┘" -ForegroundColor Yellow
    Write-Host ""
    Write-Host "  Do you want to proceed? Type 'yes' to confirm: " -ForegroundColor Yellow -NoNewline
    $confirm = Read-Host

    if ($confirm.Trim().ToLower() -ne 'yes') {
        Write-Host ""
        Write-Host "  Migration cancelled. No changes were made." -ForegroundColor Yellow
        Write-Host "  The preview file is still available at: $previewManifestPath" -ForegroundColor DarkGray
        Write-Host ""
        return
    }

    Write-Host ""
}
elseif ($migrateRows.Count -eq 0) {
    Write-Host "  No repositories to migrate. All rows are archived or skipped." -ForegroundColor Yellow
    return
}

# ─── Execute Migrations ───────────────────────────────────────────────────────

Write-Host "  ╔═══════════════════════════════════════════════════════╗" -ForegroundColor Cyan
Write-Host "  ║              Migration In Progress                    ║" -ForegroundColor Cyan
Write-Host "  ╚═══════════════════════════════════════════════════════╝" -ForegroundColor Cyan
Write-Host ""

$startTime = Get-Date
$processedCount = 0
$successCount = 0
$failCount = 0
$skippedStuckCount = 0

# Resolve timeouts from config (can be overridden per-item in the future)
$cfgTimeoutMin = if ($config.migrationDefaults.timeoutMinutes) { [int]$config.migrationDefaults.timeoutMinutes } else { 0 }
$cfgStallMin = if ($config.migrationDefaults.stallTimeoutMinutes) { [int]$config.migrationDefaults.stallTimeoutMinutes } else { 0 }

# Group migrations by Collection/Project/Repo to batch operations on same repo
$grouped = $migrateRows | Group-Object -Property { "$($_.Collection)|$($_.Project)|$($_.Repo)" }

foreach ($group in $grouped) {
    $parts = $group.Name -split '\|'
    $sourceCollection = $parts[0]
    $sourceProject = $parts[1]
    $sourceRepo = $parts[2]

    $pat = $config.collections[$sourceCollection]?.pat
    if (-not $pat) {
        $friendlyMsg = "The collection '$sourceCollection' is not in your configuration file. " +
            "Add it using the Setup Wizard (option 1 on the main menu)."
        Write-MigrationLog -Message "Collection '$sourceCollection' not found in config" -LogFile $logFile -Level WARN
        foreach ($item in $group.Group) {
            $item.Status = 'Failed'
            $item.Error = $friendlyMsg
            $failCount++
        }
        $processedCount += $group.Group.Count
        Write-Host "  ⚠ Skipping $($group.Group.Count) item(s) in '$sourceCollection' — collection not in config" -ForegroundColor Yellow
        continue
    }

    # Separate full-repo migrations from folder spin-outs in this group
    $repoItems = @($group.Group | Where-Object { $_.Action -eq 'MigrateRepo' })
    $folderItems = @($group.Group | Where-Object { $_.Action -eq 'SpinOutFolder' })

    # ── Full Repo Migration ──
    foreach ($item in $repoItems) {
        $processedCount++
        $itemStart = Get-Date
        $pctComplete = [math]::Round(($processedCount / $migrateRows.Count) * 100)
        $etaStr = ''
        if ($processedCount -gt 1) {
            $avgSec = ((Get-Date) - $startTime).TotalSeconds / ($processedCount - 1)
            $remainingSec = $avgSec * ($migrateRows.Count - $processedCount + 1)
            $eta = (Get-Date).AddSeconds($remainingSec)
            $etaStr = " | ETA: $($eta.ToString('HH:mm:ss'))"
        }
        Write-Host ""
        Write-Host "  [$processedCount / $($migrateRows.Count)] ($pctComplete%)$etaStr" -ForegroundColor Cyan
        Write-Host "    Migrating repo: $sourceCollection / $sourceProject / $sourceRepo" -ForegroundColor White
        Write-MigrationLog -Message "[$processedCount/$($migrateRows.Count)] Starting: $sourceCollection/$sourceProject/$sourceRepo" -LogFile $logFile -Level INFO

        try {
            $needsConversion = $item.RepoType -eq 'TFVC'
            $outputName = $sourceRepo -replace '[^a-zA-Z0-9_-]', '-'
            $repoTfvcPath = "`$/$sourceProject/$sourceRepo"

            if ($needsConversion) {
                if (-not (Test-TfvcPathExists -ServerUrl $config.adoServerUrl -Collection $sourceCollection -Pat $pat -TfvcPath $repoTfvcPath)) {
                    throw "TFVC path not found: $repoTfvcPath"
                }

                Write-Host "    Converting from TFVC to Git format..." -ForegroundColor DarkGray
                Write-MigrationLog -Message "Converting TFVC repo '$sourceRepo' to Git" -LogFile $logFile -Level INFO

                & "$PSScriptRoot/Convert-TfvcToGit.ps1" `
                    -ConfigPath $ConfigPath `
                    -Collection $sourceCollection `
                    -ProjectName $sourceProject `
                    -TfvcPath $repoTfvcPath `
                    -OutputRepoName $outputName `
                    -NonInteractive `
                    -TimeoutMinutes $cfgTimeoutMin `
                    -StallTimeoutMinutes $cfgStallMin
            }

            $localRepoPath = Join-Path $config.outputDirectory $outputName
            if (-not (Test-Path (Join-Path $localRepoPath '.git'))) {
                throw "Local converted repo not found at $localRepoPath. Conversion may have been skipped or failed earlier."
            }

            Write-Host "    Moving to $TargetCollection / $TargetProject..." -ForegroundColor DarkGray
            Write-MigrationLog -Message "Moving '$sourceRepo' to $TargetCollection/$TargetProject" -LogFile $logFile -Level INFO

            & "$PSScriptRoot/Move-RepoToCollection.ps1" `
                -ConfigPath $ConfigPath `
                -SourceCollection $sourceCollection `
                -SourceProject $sourceProject `
                -TfvcPath "`$/$sourceProject/$sourceRepo" `
                -TargetCollection $TargetCollection `
                -TargetProject $TargetProject `
                -TargetRepoName $outputName `
                -SkipConversion

            $itemDuration = (Get-Date) - $itemStart
            $item.Status = 'Success'
            $successCount++
            Write-Host "    ✓ Done ($($itemDuration.ToString('hh\:mm\:ss')))" -ForegroundColor Green
            Write-MigrationLog -Message "Successfully migrated '$sourceRepo' in $($itemDuration.ToString('hh\:mm\:ss'))" -LogFile $logFile -Level SUCCESS

            # Update spreadsheet column H to mark as completed
            $rowsToMark = @($item.ExcelRow)
            # Include any duplicate rows for the same repo
            $skipRows | Where-Object {
                $_.Collection -eq $item.Collection -and
                $_.Project -eq $item.Project -and
                $_.Repo -eq $item.Repo -and
                $_.Reason -match 'Duplicate row'
            } | ForEach-Object { $rowsToMark += $_.ExcelRow }

            try {
                $pkg = Open-ExcelPackage -Path $ExcelPath
                $ws = $pkg.Workbook.Worksheets[$WorksheetName]
                foreach ($excelRow in $rowsToMark) {
                    $ws.Cells[$excelRow, $recommendationCol].Value = 'completed'
                }
                Close-ExcelPackage $pkg -SaveAs $ExcelPath
                Write-MigrationLog -Message "Updated $($rowsToMark.Count) Excel row(s) Recommendation -> 'completed'" -LogFile $logFile -Level INFO
            }
            catch {
                Write-MigrationLog -Message "Could not update Excel: $($_.Exception.Message)" -LogFile $logFile -Level WARN
                Write-Host "    * Could not update spreadsheet (file may be open elsewhere)" -ForegroundColor Yellow
            }

            # Clean up local clone to free disk space
            $localClonePath = Join-Path $config.outputDirectory $outputName
            if (Test-Path $localClonePath) {
                Remove-Item -Recurse -Force $localClonePath
                Write-MigrationLog -Message "Cleaned up local clone: $localClonePath" -LogFile $logFile -Level INFO
                Write-Host "    - Cleaned up local clone" -ForegroundColor DarkGray
            }
        }
        catch {
            $errMsg = $_.Exception.Message
            $friendlyErr = Get-FriendlyError -ErrorMessage $errMsg

            # Detect path-too-long errors specifically
            if ($errMsg -match 'too long|path.*long|248 char|260 char|TF400959|PathTooLong|could not find a part of the path') {
                $item.Status = 'PathTooLong'
                $item.Error = "File paths in this repo exceed Windows' 260-character limit. " +
                    "Consider splitting by folder or renaming deeply nested TFVC paths."
                Write-Host "    ✗ PATH TOO LONG — $sourceRepo" -ForegroundColor Red
                Write-Host "      This repo contains files exceeding the Windows path limit." -ForegroundColor Yellow
            }
            elseif ($errMsg -match 'timed out|stalled') {
                $item.Status = 'TimedOut'
                $item.Error = "Operation timed out or stalled. Skipped to unblock remaining items."
                $skippedStuckCount++
                Write-Host "    ✗ TIMED OUT / STUCK — $sourceRepo (skipping to next item)" -ForegroundColor Magenta
                Write-MigrationLog -Message "SKIPPED (timeout/stall): '$sourceRepo' — moving on to next item" -LogFile $logFile -Level WARN
            }
            else {
                $item.Status = 'Failed'
                $item.Error = $friendlyErr
                Write-Host "    ✗ Failed: $friendlyErr" -ForegroundColor Red
            }

            $failCount++
            Write-MigrationLog -Message "Failed to migrate '$sourceRepo': $errMsg" -LogFile $logFile -Level ERROR
        }
    }

    # ── Folder Spin-Outs ──
    # Clone each folder independently from its own TFVC path rather than cloning
    # the entire repo and filtering. This avoids checking out deeply nested files
    # from sibling folders (which can cause PathTooLong failures on Windows).
    if ($folderItems.Count -gt 0) {

        foreach ($item in $folderItems) {
            $processedCount++
            $itemStart = Get-Date
            $pctComplete = [math]::Round(($processedCount / $migrateRows.Count) * 100)
            $etaStr = ''
            if ($processedCount -gt 1) {
                $avgSec = ((Get-Date) - $startTime).TotalSeconds / ($processedCount - 1)
                $remainingSec = $avgSec * ($migrateRows.Count - $processedCount + 1)
                $eta = (Get-Date).AddSeconds($remainingSec)
                $etaStr = " | ETA: $($eta.ToString('HH:mm:ss'))"
            }
            $folderName = $item.Folder
            $folderOutputName = $item.NewRepoName
            $folderOutputPath = Join-Path $config.outputDirectory $folderOutputName
            $folderTfvcPath = Resolve-TfvcFolderPath -SourceProject $sourceProject -SourceRepo $sourceRepo -FolderName $folderName

            if (-not (Test-TfvcPathExists -ServerUrl $config.adoServerUrl -Collection $sourceCollection -Pat $pat -TfvcPath $folderTfvcPath)) {
                $item.Status = 'Failed'
                $item.Error = "TFVC path not found: $folderTfvcPath"
                $failCount++
                Write-Host ""
                Write-Host "  [$processedCount / $($migrateRows.Count)] ($pctComplete%)$etaStr" -ForegroundColor Cyan
                Write-Host "    Extracting folder: $folderTfvcPath → '$folderOutputName'" -ForegroundColor White
                Write-Host "    ✗ Failed: TFVC path not found" -ForegroundColor Red
                Write-MigrationLog -Message "Failed to extract '$folderName': TFVC path not found ($folderTfvcPath)" -LogFile $logFile -Level ERROR
                continue
            }

            Write-Host ""
            Write-Host "  [$processedCount / $($migrateRows.Count)] ($pctComplete%)$etaStr" -ForegroundColor Cyan
            Write-Host "    Extracting folder: $folderTfvcPath → '$folderOutputName'" -ForegroundColor White
            Write-MigrationLog -Message "[$processedCount/$($migrateRows.Count)] Starting folder: $folderTfvcPath" -LogFile $logFile -Level INFO

            # Pre-clean output directory
            if (Test-Path $folderOutputPath) {
                Remove-Item -Recurse -Force $folderOutputPath
            }

            try {
                # Clone directly from the folder's TFVC path — only checks out files in this folder
                & "$PSScriptRoot/Convert-TfvcToGit.ps1" `
                    -ConfigPath $ConfigPath `
                    -Collection $sourceCollection `
                    -ProjectName $sourceProject `
                    -TfvcPath $folderTfvcPath `
                    -OutputRepoName $folderOutputName `
                    -NonInteractive `
                    -TimeoutMinutes $cfgTimeoutMin `
                    -StallTimeoutMinutes $cfgStallMin

                Write-Host "    ✓ Extracted '$folderName'" -ForegroundColor Green
                Write-MigrationLog -Message "Successfully extracted '$folderName' as '$folderOutputName'" -LogFile $logFile -Level SUCCESS
            }
            catch {
                $errMsg = $_.Exception.Message
                $friendlyErr = Get-FriendlyError -ErrorMessage $errMsg

                # Detect path-too-long errors specifically
                if ($errMsg -match 'too long|path.*long|248 char|260 char|TF400959|PathTooLong|could not find a part of the path') {
                    $item.Status = 'PathTooLong'
                    $item.Error = "File paths in this folder exceed Windows' 260-character limit. " +
                        "This folder must be migrated manually or the deeply nested files renamed in TFVC first."
                    Write-Host "    ✗ PATH TOO LONG — $folderName" -ForegroundColor Red
                    Write-Host "      This folder contains files with paths exceeding the Windows limit." -ForegroundColor Yellow
                    Write-Host "      Mark this item as 'Skipped' in the manifest to exclude it from future runs." -ForegroundColor Yellow
                }
                elseif ($errMsg -match 'timed out|stalled') {
                    $item.Status = 'TimedOut'
                    $item.Error = "Operation timed out or stalled. Skipped to unblock remaining items."
                    $skippedStuckCount++
                    Write-Host "    ✗ TIMED OUT / STUCK — $folderName (skipping to next item)" -ForegroundColor Magenta
                    Write-MigrationLog -Message "SKIPPED (timeout/stall): folder '$folderName' — moving on" -LogFile $logFile -Level WARN
                }
                else {
                    $item.Status = 'Failed'
                    $item.Error = $friendlyErr
                    Write-Host "    ✗ Failed: $friendlyErr" -ForegroundColor Red
                }

                $failCount++
                Write-MigrationLog -Message "Failed to extract '$folderName': $errMsg" -LogFile $logFile -Level ERROR
                continue
            }

            # Move to target collection/project
            if (-not (Test-Path (Join-Path $folderOutputPath '.git'))) {
                $item.Status = 'Failed'
                $item.Error = "Local converted repo not found at $folderOutputPath. Conversion may have been skipped or failed earlier."
                $failCount++
                Write-Host "    ✗ Failed: $($item.Error)" -ForegroundColor Red
                Write-MigrationLog -Message "Failed to move '$folderName': $($item.Error)" -LogFile $logFile -Level ERROR
                continue
            }

            Write-Host "    Moving '$folderOutputName' to $TargetCollection / $TargetProject..." -ForegroundColor White

            try {
                & "$PSScriptRoot/Move-RepoToCollection.ps1" `
                    -ConfigPath $ConfigPath `
                    -SourceCollection $sourceCollection `
                    -SourceProject $sourceProject `
                    -TfvcPath $folderTfvcPath `
                    -TargetCollection $TargetCollection `
                    -TargetProject $TargetProject `
                    -TargetRepoName $folderOutputName `
                    -SkipConversion

                $itemDuration = (Get-Date) - $itemStart
                $item.Status = 'Success'
                $successCount++
                Write-Host "    ✓ Done ($($itemDuration.ToString('hh\:mm\:ss')))" -ForegroundColor Green
                Write-MigrationLog -Message "Successfully moved '$folderOutputName' to $TargetCollection/$TargetProject in $($itemDuration.ToString('hh\:mm\:ss'))" -LogFile $logFile -Level SUCCESS

                # Update spreadsheet column H to mark as completed
                try {
                    $pkg = Open-ExcelPackage -Path $ExcelPath
                    $ws = $pkg.Workbook.Worksheets[$WorksheetName]
                    $ws.Cells[$item.ExcelRow, $recommendationCol].Value = 'completed'
                    Close-ExcelPackage $pkg -SaveAs $ExcelPath
                    Write-MigrationLog -Message "Updated Excel row $($item.ExcelRow) Recommendation -> 'completed'" -LogFile $logFile -Level INFO
                }
                catch {
                    Write-MigrationLog -Message "Could not update Excel row $($item.ExcelRow): $($_.Exception.Message)" -LogFile $logFile -Level WARN
                    Write-Host "    * Could not update spreadsheet (file may be open elsewhere)" -ForegroundColor Yellow
                }

                # Clean up local clone to free disk space
                if (Test-Path $folderOutputPath) {
                    Remove-Item -Recurse -Force $folderOutputPath
                    Write-MigrationLog -Message "Cleaned up local clone: $folderOutputPath" -LogFile $logFile -Level INFO
                    Write-Host "    - Cleaned up local clone" -ForegroundColor DarkGray
                }
            }
            catch {
                $item.Status = 'Failed'
                $item.Error = Get-FriendlyError -ErrorMessage $_.Exception.Message
                $failCount++
                Write-Host "    ✗ Failed: $($item.Error)" -ForegroundColor Red
                Write-MigrationLog -Message "Failed to move '$folderName': $($_.Exception.Message)" -LogFile $logFile -Level ERROR
            }
        }
    }

    # ── Git Mirror (existing Git repos → GAMS-GIT-Repos collection) ──
    # For rows with RepoType=Git + Recommendation=Migrate, perform a
    # `git clone --mirror` + `git push --mirror` from the source collection
    # to the configured Git target collection. Target project defaults to
    # the source project name and is auto-created if missing.
    $gitItems = @($group.Group | Where-Object { $_.Action -eq 'MirrorGitRepo' })

    foreach ($item in $gitItems) {
        $processedCount++
        $itemStart = Get-Date
        $pctComplete = [math]::Round(($processedCount / $migrateRows.Count) * 100)
        $etaStr = ''
        if ($processedCount -gt 1) {
            $avgSec = ((Get-Date) - $startTime).TotalSeconds / ($processedCount - 1)
            $remainingSec = $avgSec * ($migrateRows.Count - $processedCount + 1)
            $eta = (Get-Date).AddSeconds($remainingSec)
            $etaStr = " | ETA: $($eta.ToString('HH:mm:ss'))"
        }

        $outputName = $item.NewRepoName
        $targetProjectName = if ($GitTargetProject) { $GitTargetProject } else { $sourceProject }

        Write-Host ""
        Write-Host "  [$processedCount / $($migrateRows.Count)] ($pctComplete%)$etaStr" -ForegroundColor Cyan
        Write-Host "    Mirroring Git repo: $sourceCollection / $sourceProject / $sourceRepo" -ForegroundColor White
        Write-Host "                    →  $GitTargetCollection / $targetProjectName / $outputName" -ForegroundColor DarkGray
        Write-MigrationLog -Message "[$processedCount/$($migrateRows.Count)] Mirror start: $sourceCollection/$sourceProject/$sourceRepo → $GitTargetCollection/$targetProjectName/$outputName" -LogFile $logFile -Level INFO

        $localBare = $null
        try {
            # Resolve source + target endpoints (each may live on a different
            # ADO server thanks to per-collection serverUrl overrides).
            $sourceServerUrl = Get-CollectionServerUrl -Config $config -Collection $sourceCollection
            $sourcePat       = $pat   # already resolved above for this $group
            $targetServerUrl = Get-CollectionServerUrl -Config $config -Collection $GitTargetCollection
            $targetPat       = $config.collections[$GitTargetCollection].pat
            if (-not $targetPat -or $targetPat -eq 'YOUR_PAT_HERE') {
                throw "PAT for Git target collection '$GitTargetCollection' is not set in config."
            }

            # 1. Resolve the source Git repo (we need its remoteUrl).
            $srcRepo = Get-AdoGitRepository -ServerUrl $sourceServerUrl `
                -Collection $sourceCollection -ProjectIdOrName $sourceProject `
                -RepoName $sourceRepo -Pat $sourcePat
            if (-not $srcRepo) {
                throw "Source Git repo '$sourceRepo' not found in $sourceCollection/$sourceProject."
            }

            # 2. Ensure the target project exists (create if missing).
            $tgtProj = Get-AdoTeamProject -ServerUrl $targetServerUrl `
                -Collection $GitTargetCollection -Name $targetProjectName -Pat $targetPat
            if (-not $tgtProj) {
                Write-Host "    Creating target project '$targetProjectName' in $GitTargetCollection..." -ForegroundColor DarkGray
                Write-MigrationLog -Message "Creating target project '$targetProjectName' in $GitTargetCollection" -LogFile $logFile -Level INFO
                $tgtProj = New-AdoTeamProject -ServerUrl $targetServerUrl `
                    -Collection $GitTargetCollection -Pat $targetPat `
                    -Name $targetProjectName `
                    -Description "Mirrored from $sourceCollection/$sourceProject by Invoke-ExcelMigration.ps1" `
                    -LogFile $logFile
            }

            # 3. Ensure the target repo exists (create if missing; tolerate 409).
            $tgtRepo = Get-AdoGitRepository -ServerUrl $targetServerUrl `
                -Collection $GitTargetCollection -ProjectIdOrName $tgtProj.id `
                -RepoName $outputName -Pat $targetPat
            if (-not $tgtRepo) {
                try {
                    $tgtRepo = New-AdoGitRepository -ServerUrl $targetServerUrl `
                        -Collection $GitTargetCollection -ProjectId $tgtProj.id `
                        -Name $outputName -Pat $targetPat
                    Write-MigrationLog -Message "Created target repo '$outputName' in $GitTargetCollection/$targetProjectName" -LogFile $logFile -Level SUCCESS
                }
                catch {
                    if ($_.Exception.Message -match '409|already exists|TF400898') {
                        $tgtRepo = Get-AdoGitRepository -ServerUrl $targetServerUrl `
                            -Collection $GitTargetCollection -ProjectIdOrName $tgtProj.id `
                            -RepoName $outputName -Pat $targetPat
                        if (-not $tgtRepo) { throw }
                    }
                    else { throw }
                }
            }
            else {
                Write-MigrationLog -Message "Target repo '$outputName' already exists in $GitTargetCollection/$targetProjectName — mirroring will overwrite refs" -LogFile $logFile -Level WARN
            }

            # 4. Clone --mirror locally, then push --mirror to the target.
            $localBare = Join-Path $config.outputDirectory ("{0}.git" -f $outputName)
            if (Test-Path $localBare) {
                Remove-Item -Recurse -Force $localBare -ErrorAction SilentlyContinue
            }

            $srcAuth = Add-PatToGitUrl -Url $srcRepo.remoteUrl -Pat $sourcePat
            $tgtAuth = Add-PatToGitUrl -Url $tgtRepo.remoteUrl -Pat $targetPat

            Invoke-GitMirror `
                -Arguments "clone --mirror `"$srcAuth`" `"$localBare`"" `
                -LogFile $logFile `
                -StatusLabel 'Cloning (mirror)' `
                -TimeoutMinutes $cfgTimeoutMin `
                -StallTimeoutMinutes $cfgStallMin | Out-Null

            Invoke-GitMirror `
                -Arguments "push --mirror `"$tgtAuth`"" `
                -WorkingDirectory $localBare `
                -LogFile $logFile `
                -StatusLabel 'Pushing (mirror)' `
                -TimeoutMinutes $cfgTimeoutMin `
                -StallTimeoutMinutes $cfgStallMin | Out-Null

            $itemDuration = (Get-Date) - $itemStart
            $item.Status = 'Success'
            $successCount++
            Write-Host "    ✓ Done ($($itemDuration.ToString('hh\:mm\:ss')))" -ForegroundColor Green
            Write-MigrationLog -Message "Successfully mirrored '$sourceRepo' in $($itemDuration.ToString('hh\:mm\:ss'))" -LogFile $logFile -Level SUCCESS

            # Update spreadsheet column H to mark as completed
            $rowsToMark = @($item.ExcelRow)
            $skipRows | Where-Object {
                $_.Collection -eq $item.Collection -and
                $_.Project -eq $item.Project -and
                $_.Repo -eq $item.Repo -and
                $_.Reason -match 'Duplicate row'
            } | ForEach-Object { $rowsToMark += $_.ExcelRow }

            try {
                $pkg = Open-ExcelPackage -Path $ExcelPath
                $ws = $pkg.Workbook.Worksheets[$WorksheetName]
                foreach ($excelRow in $rowsToMark) {
                    $ws.Cells[$excelRow, $recommendationCol].Value = 'completed'
                }
                Close-ExcelPackage $pkg -SaveAs $ExcelPath
                Write-MigrationLog -Message "Updated $($rowsToMark.Count) Excel row(s) Recommendation -> 'completed'" -LogFile $logFile -Level INFO
            }
            catch {
                Write-MigrationLog -Message "Could not update Excel: $($_.Exception.Message)" -LogFile $logFile -Level WARN
                Write-Host "    * Could not update spreadsheet (file may be open elsewhere)" -ForegroundColor Yellow
            }

            # Clean up local bare clone to free disk space
            if (Test-Path $localBare) {
                Remove-Item -Recurse -Force $localBare -ErrorAction SilentlyContinue
                Write-MigrationLog -Message "Cleaned up local bare clone: $localBare" -LogFile $logFile -Level INFO
                Write-Host "    - Cleaned up local bare clone" -ForegroundColor DarkGray
            }
        }
        catch {
            $errMsg = $_.Exception.Message
            $friendlyErr = Get-FriendlyError -ErrorMessage $errMsg

            if ($errMsg -match 'too long|PathTooLong|260 char') {
                $item.Status = 'PathTooLong'
                $item.Error = "Local bare-clone path exceeded Windows' limit. Try a shorter outputDirectory."
                Write-Host "    ✗ PATH TOO LONG — $sourceRepo" -ForegroundColor Red
            }
            elseif ($errMsg -match 'TIMED OUT|timed out|STALL|stalled') {
                $item.Status = 'TimedOut'
                $item.Error = "Mirror operation timed out or stalled. Skipped to unblock remaining items."
                $skippedStuckCount++
                Write-Host "    ✗ TIMED OUT / STUCK — $sourceRepo (skipping to next item)" -ForegroundColor Magenta
                Write-MigrationLog -Message "SKIPPED (timeout/stall): mirror '$sourceRepo' — moving on" -LogFile $logFile -Level WARN
            }
            else {
                $item.Status = 'Failed'
                $item.Error = $friendlyErr
                Write-Host "    ✗ Failed: $friendlyErr" -ForegroundColor Red
            }

            $failCount++
            Write-MigrationLog -Message "Failed to mirror '$sourceRepo': $errMsg" -LogFile $logFile -Level ERROR

            # Best-effort cleanup so a failed mirror doesn't leave a half-baked bare clone behind.
            if ($localBare -and (Test-Path $localBare)) {
                Remove-Item -Recurse -Force $localBare -ErrorAction SilentlyContinue
            }
        }
    }
}

# ─── Final Report ─────────────────────────────────────────────────────────────

$totalDuration = (Get-Date) - $startTime

# Build comprehensive manifest
$allActions = [System.Collections.ArrayList]::new()
foreach ($r in $migrateRows) {
    $r | Add-Member -NotePropertyName Destination -NotePropertyValue "$TargetCollection/$TargetProject" -Force
    $r | Add-Member -NotePropertyName ConvertedToGit -NotePropertyValue ($r.RepoType -eq 'TFVC') -Force
    $r | Add-Member -NotePropertyName SpunOut -NotePropertyValue ($r.Action -eq 'SpinOutFolder') -Force
    [void]$allActions.Add($r)
}
foreach ($r in $archiveRows) {
    $r | Add-Member -NotePropertyName Destination -NotePropertyValue 'N/A (Archived — not migrated)' -Force
    $r | Add-Member -NotePropertyName ConvertedToGit -NotePropertyValue $false -Force
    $r | Add-Member -NotePropertyName SpunOut -NotePropertyValue $false -Force
    [void]$allActions.Add($r)
}
foreach ($r in $skipRows) {
    $r | Add-Member -NotePropertyName Destination -NotePropertyValue 'N/A (Skipped)' -Force
    $r | Add-Member -NotePropertyName ConvertedToGit -NotePropertyValue $false -Force
    $r | Add-Member -NotePropertyName SpunOut -NotePropertyValue $false -Force
    [void]$allActions.Add($r)
}

# Write CSV manifest
$manifestPath = Join-Path $config.outputDirectory "excel-migration-manifest-$(Get-Date -Format 'yyyyMMdd-HHmmss').csv"
$allActions | Select-Object RowNo, Collection, Project, Repo, Folder, NewRepoName, RepoType, RepoOrFolder, `
    Recommendation, Action, Reason, Status, Error, Destination, ConvertedToGit, SpunOut |
    Export-Csv -Path $manifestPath -NoTypeInformation -Encoding utf8

# Write JSON report
$reportPath = Join-Path $config.outputDirectory "excel-migration-report-$(Get-Date -Format 'yyyyMMdd-HHmmss').json"
$report = @{
    startTime           = $startTime.ToString('o')
    endTime             = (Get-Date).ToString('o')
    duration            = $totalDuration.ToString('hh\:mm\:ss')
    excelFile           = $ExcelPath
    worksheet           = $WorksheetName
    target              = "$TargetCollection/$TargetProject"
    totalRows           = $excelData.Count
    migrateCount        = $migrateRows.Count
    archiveCount        = $archiveRows.Count
    skipCount           = $skipRows.Count
    successCount        = $successCount
    failCount           = $failCount
    pathTooLongCount    = $pathTooLongCount
    timedOutCount       = $skippedStuckCount
}
$report | ConvertTo-Json -Depth 5 | Out-File $reportPath -Encoding utf8

Write-MigrationLog -Message "Migration manifest: $manifestPath" -LogFile $logFile -Level SUCCESS
Write-MigrationLog -Message "Report: $reportPath" -LogFile $logFile -Level INFO

# Display results
Write-Host ""
Write-Host ""
Write-Host "  ╔═══════════════════════════════════════════════════════╗" -ForegroundColor Cyan
Write-Host "  ║              Migration Complete                       ║" -ForegroundColor Cyan
Write-Host "  ╚═══════════════════════════════════════════════════════╝" -ForegroundColor Cyan
Write-Host ""
Write-Host "  Time elapsed:  $($totalDuration.ToString('hh\:mm\:ss'))" -ForegroundColor White
Write-Host ""

$pathTooLongCount = @($migrateRows | Where-Object { $_.Status -eq 'PathTooLong' }).Count
$timedOutCount = @($migrateRows | Where-Object { $_.Status -eq 'TimedOut' }).Count
$otherFailCount = $failCount - $pathTooLongCount - $timedOutCount

Write-Host "  Results:" -ForegroundColor White
Write-Host "    ✓ Successful:    $successCount" -ForegroundColor Green
if ($timedOutCount -gt 0) {
    Write-Host "    ⏱ Timed out:    $timedOutCount (stuck or exceeded timeout — skipped)" -ForegroundColor Magenta
}
if ($pathTooLongCount -gt 0) {
    Write-Host "    ✗ Path too long: $pathTooLongCount (file paths exceed Windows limit)" -ForegroundColor Magenta
}
if ($otherFailCount -gt 0) {
    Write-Host "    ✗ Failed:        $otherFailCount" -ForegroundColor Red
}
Write-Host "    ○ Archived:      $($archiveRows.Count) (not migrated)" -ForegroundColor Yellow
Write-Host "    ○ Skipped:       $($skipRows.Count)" -ForegroundColor DarkGray
Write-Host ""

# Show Timed-out items
if ($timedOutCount -gt 0) {
    Write-Host "  ┌─ Items that timed out / appeared stuck ───────────────┐" -ForegroundColor Magenta
    foreach ($f in ($migrateRows | Where-Object { $_.Status -eq 'TimedOut' })) {
        $label = "$($f.Collection)/$($f.Project)/$($f.Repo)"
        if ($f.Folder) { $label += "/$($f.Folder)" }
        Write-Host "  │  ⏱ $label" -ForegroundColor Magenta
    }
    Write-Host "  │" -ForegroundColor Magenta
    Write-Host "  │  These items exceeded the timeout or stalled with no" -ForegroundColor Magenta
    Write-Host "  │  progress. They were skipped to unblock the batch." -ForegroundColor Magenta
    Write-Host "  │  You can retry them individually or increase the" -ForegroundColor Magenta
    Write-Host "  │  timeoutMinutes / stallTimeoutMinutes in your config." -ForegroundColor Magenta
    Write-Host "  └───────────────────────────────────────────────────────┘" -ForegroundColor Magenta
    Write-Host ""
}

# Show PathTooLong items with specific guidance
if ($pathTooLongCount -gt 0) {
    Write-Host "  ┌─ Items with paths too long ───────────────────────────┐" -ForegroundColor Magenta
    foreach ($f in ($migrateRows | Where-Object { $_.Status -eq 'PathTooLong' })) {
        $label = "$($f.Collection)/$($f.Project)/$($f.Repo)"
        if ($f.Folder) { $label += "/$($f.Folder)" }
        Write-Host "  │  ✗ $label" -ForegroundColor Magenta
    }
    Write-Host "  │" -ForegroundColor Magenta
    Write-Host "  │  These items contain deeply nested file paths that" -ForegroundColor Magenta
    Write-Host "  │  exceed Windows' 260-character limit. Options:" -ForegroundColor Magenta
    Write-Host "  │    1. Rename the deeply nested folders in TFVC" -ForegroundColor Yellow
    Write-Host "  │    2. Migrate them manually via a different method" -ForegroundColor Yellow
    Write-Host "  │    3. Mark as 'Skipped' in the manifest CSV to" -ForegroundColor Yellow
    Write-Host "  │       exclude from future runs" -ForegroundColor Yellow
    Write-Host "  └───────────────────────────────────────────────────────┘" -ForegroundColor Magenta
    Write-Host ""
}

if ($otherFailCount -gt 0) {
    Write-Host "  ┌─ Items that failed ───────────────────────────────────┐" -ForegroundColor Red
    foreach ($f in ($migrateRows | Where-Object { $_.Status -eq 'Failed' }) | Select-Object -First 15) {
        $label = "$($f.Collection)/$($f.Project)/$($f.Repo)"
        if ($f.Folder) { $label += "/$($f.Folder)" }
        Write-Host "  │  ✗ $label" -ForegroundColor Red
        Write-Host "  │    $($f.Error)" -ForegroundColor DarkGray
    }
    if ($otherFailCount -gt 15) {
        Write-Host "  │  ... and $($otherFailCount - 15) more (see manifest for full list)" -ForegroundColor Red
    }
    Write-Host "  └───────────────────────────────────────────────────────┘" -ForegroundColor Red
    Write-Host ""
}

if ($failCount -gt 0) {
    Write-Host "  To resume and skip failed items:" -ForegroundColor Cyan
    Write-Host "    1. Open the manifest CSV: $manifestPath" -ForegroundColor DarkGray
    Write-Host "    2. Change Status of items to skip from 'Failed'/'PathTooLong' to 'Skipped'" -ForegroundColor DarkGray
    Write-Host "    3. Re-run with -ResumeManifest `"$manifestPath`"" -ForegroundColor DarkGray
    Write-Host ""
}

Write-Host "  Output files:" -ForegroundColor White
Write-Host "    Manifest (CSV):  $manifestPath" -ForegroundColor DarkGray
Write-Host "    Report (JSON):   $reportPath" -ForegroundColor DarkGray
Write-Host "    Log file:        $logFile" -ForegroundColor DarkGray
Write-Host ""
Write-Host "  The manifest CSV can be opened in Excel for a complete" -ForegroundColor DarkGray
Write-Host "  record of every repository and what happened to it." -ForegroundColor DarkGray
Write-Host ""
