#Requires -Version 7.0
<#
.SYNOPSIS
    Processes the MDR-4ADO-AllProjects Excel file to batch migrate/split/skip TFVC repos.

.DESCRIPTION
    Reads the "GAMS-Repos-App-Folder level" worksheet from MDR-4ADO-AllProjects.xlsx
    and processes each row based on the Recommendation and Repo/Folder columns:

    - Recommendation = "Archive*" → Skip (no migration)
    - Repo or Folder = "Repo" + Recommendation = "Migrate*" → Convert entire repo from TFVC to Git, move to target
    - Repo or Folder = "Folder" + Recommendation = "Migrate*" → Spin out folder from parent repo into new Git repo, move to target

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
    Target ADO collection to move migrated repos into. Required.

.PARAMETER SourceServerUrl
    Optional source ADO server URL override. If omitted, uses sourceAdoServerUrl
    from config, then falls back to legacy adoServerUrl.

.PARAMETER TargetServerUrl
    Optional target ADO server URL override. If omitted, uses targetAdoServerUrl
    from config, then falls back to legacy adoServerUrl.

.PARAMETER TargetProject
    Target ADO project to move migrated repos into. Required.

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
    ./Invoke-ExcelMigration.ps1 -ConfigPath ./config/migration-config.json `
        -ExcelPath ./excel-docs/MDR-4ADO-AllProjects.xlsx `
        -TargetCollection "ModernApps" -TargetProject "Platform"
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory)]
    [string]$ConfigPath,

    [string]$ExcelPath,

    [string]$WorksheetName = 'GAMS-Repos-App-Folder level',

    [string]$SourceServerUrl,

    [string]$TargetCollection,

    [string]$TargetProject,

    [string]$TargetServerUrl,

    [string]$ResumeManifest,

    [switch]$Interactive
)

$ErrorActionPreference = 'Stop'
Import-Module "$PSScriptRoot/modules/AdoTfvcMigrator.psm1" -Force

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
if (-not $SourceServerUrl) {
    $SourceServerUrl = Get-ConfigAdoServerUrl -Config $config -Role Source
}
if (-not $TargetServerUrl) {
    $TargetServerUrl = Get-ConfigAdoServerUrl -Config $config -Role Target
}
Write-MigrationLog -Message "Source ADO URL: $SourceServerUrl" -LogFile $logFile -Level INFO
Write-MigrationLog -Message "Target ADO URL: $TargetServerUrl" -LogFile $logFile -Level INFO

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

# ─── Step 2: Choose the Target ────────────────────────────────────────────────

if ($Interactive) {
    Write-Host "  ─────────────────────────────────────────────────────" -ForegroundColor DarkGray
    Write-Host "  Step 2 of 4: Where should migrated repos go?" -ForegroundColor Cyan
    Write-Host "  ─────────────────────────────────────────────────────" -ForegroundColor DarkGray
    Write-Host ""
    Write-Host "  Repos marked for migration will be moved to a target" -ForegroundColor DarkGray
    Write-Host "  collection and project. Select them below." -ForegroundColor DarkGray
    Write-Host ""
}

if (-not $TargetCollection -and $Interactive) {
    $TargetCollection = Select-AdoCollection -Config $config -Prompt 'Select the target collection'
    if (-not $TargetCollection) {
        Write-Host "  Cancelled." -ForegroundColor Yellow
        return
    }
}

if (-not $TargetProject -and $Interactive) {
    Write-Host ""
    Write-Host "  Enter the target project name: " -ForegroundColor Yellow -NoNewline
    $TargetProject = Read-Host
    if (-not $TargetProject) {
        Write-Host "  Cancelled." -ForegroundColor Yellow
        return
    }
}

if (-not $TargetCollection -or -not $TargetProject) {
    Write-Host "  Target collection and project are required." -ForegroundColor Red
    return
}

# ─── Step 3: Read & Analyze the Spreadsheet ───────────────────────────────────

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

    # Exclude BuildProcessTemplates
    if ($folderName -eq 'BuildProcessTemplates') {
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
        if ($repoOrFolder -eq 'Folder') {
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
                $_.Action -eq 'MigrateRepo' -and
                "$($_.Collection)|$($_.Project)|$($_.Repo)" -eq $repoKey
            }
            if ($alreadyAdded) {
                $entry.Action = 'Skip'
                $entry.Reason = "Duplicate row — repo '$repoName' is already queued for full migration"
                $entry.Status = 'Skipped'
                [void]$skipRows.Add($entry)
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
# ─── Show the Preview ─────────────────────────────────────────────────────────

$repoMigrations = @($migrateRows | Where-Object { $_.Action -eq 'MigrateRepo' })
$folderSpinouts = @($migrateRows | Where-Object { $_.Action -eq 'SpinOutFolder' })
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
    Write-Host "        └─ $($folderSpinouts.Count) folder(s) will be extracted into standalone repos" -ForegroundColor DarkGray
}
if ($tfvcConversions.Count -gt 0) {
    Write-Host "        ($($tfvcConversions.Count) of these need to be converted from TFVC to Git format)" -ForegroundColor DarkGray
}
Write-Host "    ● $($archiveRows.Count) repositories are marked for archive (will be skipped)" -ForegroundColor Yellow
Write-Host "    ● $($skipRows.Count) rows will be skipped (system folders, missing data, etc.)" -ForegroundColor DarkGray
Write-Host ""
Write-Host "    Destination: $TargetCollection / $TargetProject" -ForegroundColor White
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
    Write-Host "  │    • Move them to: $($TargetCollection)/$($TargetProject)".PadRight(37) + "│" -ForegroundColor Yellow
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

            if ($needsConversion) {
                Write-Host "    Converting from TFVC to Git format..." -ForegroundColor DarkGray
                Write-MigrationLog -Message "Converting TFVC repo '$sourceRepo' to Git" -LogFile $logFile -Level INFO

                & "$PSScriptRoot/Convert-TfvcToGit.ps1" `
                    -ConfigPath $ConfigPath `
                    -ServerUrl $SourceServerUrl `
                    -Collection $sourceCollection `
                    -ProjectName $sourceProject `
                    -TfvcPath "`$/$sourceProject" `
                    -OutputRepoName $outputName `
                    -NonInteractive `
                    -TimeoutMinutes $cfgTimeoutMin `
                    -StallTimeoutMinutes $cfgStallMin
            }

            Write-Host "    Moving to $TargetCollection / $TargetProject..." -ForegroundColor DarkGray
            Write-MigrationLog -Message "Moving '$sourceRepo' to $TargetCollection/$TargetProject" -LogFile $logFile -Level INFO

            & "$PSScriptRoot/Move-RepoToCollection.ps1" `
                -ConfigPath $ConfigPath `
                -SourceCollection $sourceCollection `
                -SourceProject $sourceProject `
                -TfvcPath "`$/$sourceProject/$sourceRepo" `
                -SourceServerUrl $SourceServerUrl `
                -TargetCollection $TargetCollection `
                -TargetProject $TargetProject `
                -TargetRepoName $outputName `
                -TargetServerUrl $TargetServerUrl `
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
            $folderTfvcPath = "`$/$sourceProject/$folderName"

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
                    -ServerUrl $SourceServerUrl `
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
            Write-Host "    Moving '$folderOutputName' to $TargetCollection / $TargetProject..." -ForegroundColor White

            try {
                & "$PSScriptRoot/Move-RepoToCollection.ps1" `
                    -ConfigPath $ConfigPath `
                    -SourceCollection $sourceCollection `
                    -SourceProject $sourceProject `
                    -TfvcPath $folderTfvcPath `
                    -SourceServerUrl $SourceServerUrl `
                    -TargetCollection $TargetCollection `
                    -TargetProject $TargetProject `
                    -TargetRepoName $folderOutputName `
                    -TargetServerUrl $TargetServerUrl `
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
