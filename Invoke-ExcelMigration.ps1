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

    Produces a comprehensive manifest of all actions taken.

.PARAMETER ConfigPath
    Path to migration-config.json.

.PARAMETER ExcelPath
    Path to the MDR-4ADO-AllProjects.xlsx file.

.PARAMETER WorksheetName
    Worksheet to read. Default: "GAMS-Repos-App-Folder level".

.PARAMETER TargetCollection
    Target ADO collection to move migrated repos into. Required.

.PARAMETER TargetProject
    Target ADO project to move migrated repos into. Required.

.PARAMETER Interactive
    Launch interactive mode — step-by-step guided flow with preview and confirmation.

.EXAMPLE
    # Interactive mode (recommended)
    ./Invoke-ExcelMigration.ps1 -ConfigPath ./config/migration-config.json -Interactive

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

    [string]$TargetCollection,

    [string]$TargetProject,

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

foreach ($row in $excelData) {
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

    # Validate: warn if collection is not in config
    if (-not $config.collections[$collection]) {
        [void]$dataWarnings.Add("Row $($row.No): Collection '$collection' is not in your config file — this row will fail if migrated")
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
        Write-Host ""
        Write-Host "  Working on $processedCount of $($migrateRows.Count)..." -ForegroundColor Cyan
        Write-Host "    Migrating repo: $sourceCollection / $sourceProject / $sourceRepo" -ForegroundColor White

        try {
            $needsConversion = $item.RepoType -eq 'TFVC'
            $outputName = $sourceRepo -replace '[^a-zA-Z0-9_-]', '-'

            if ($needsConversion) {
                Write-Host "    Converting from TFVC to Git format..." -ForegroundColor DarkGray
                Write-MigrationLog -Message "Converting TFVC repo '$sourceRepo' to Git" -LogFile $logFile -Level INFO

                & "$PSScriptRoot/Convert-TfvcToGit.ps1" `
                    -ConfigPath $ConfigPath `
                    -Collection $sourceCollection `
                    -ProjectName $sourceProject `
                    -TfvcPath "`$/$sourceProject" `
                    -OutputRepoName $outputName
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
                -TargetRepoName $outputName

            $item.Status = 'Success'
            $successCount++
            Write-Host "    ✓ Done" -ForegroundColor Green
            Write-MigrationLog -Message "Successfully migrated '$sourceRepo'" -LogFile $logFile -Level SUCCESS
        }
        catch {
            $item.Status = 'Failed'
            $item.Error = Get-FriendlyError -ErrorMessage $_.Exception.Message
            $failCount++
            Write-Host "    ✗ Failed: $($item.Error)" -ForegroundColor Red
            Write-MigrationLog -Message "Failed to migrate '$sourceRepo': $($_.Exception.Message)" -LogFile $logFile -Level ERROR
        }
    }

    # ── Folder Spin-Outs ──
    if ($folderItems.Count -gt 0) {
        $parentRepoType = $folderItems[0].RepoType

        if ($parentRepoType -eq 'TFVC') {
            $parentOutputName = $sourceRepo -replace '[^a-zA-Z0-9_-]', '-'
            $parentRepoPath = Join-Path $config.outputDirectory $parentOutputName

            if (-not (Test-Path $parentRepoPath)) {
                Write-Host ""
                Write-Host "    Converting parent repo '$sourceRepo' to Git (needed for folder extraction)..." -ForegroundColor DarkGray

                try {
                    & "$PSScriptRoot/Convert-TfvcToGit.ps1" `
                        -ConfigPath $ConfigPath `
                        -Collection $sourceCollection `
                        -ProjectName $sourceProject `
                        -TfvcPath "`$/$sourceProject" `
                        -OutputRepoName $parentOutputName

                    Write-Host "    ✓ Parent repo converted" -ForegroundColor Green
                }
                catch {
                    $friendlyMsg = Get-FriendlyError -ErrorMessage $_.Exception.Message
                    Write-Host "    ✗ Could not convert parent repo: $friendlyMsg" -ForegroundColor Red
                    Write-MigrationLog -Message "Failed to convert parent repo '$sourceRepo': $($_.Exception.Message)" -LogFile $logFile -Level ERROR
                    foreach ($fi in $folderItems) {
                        $fi.Status = 'Failed'
                        $fi.Error = "Could not convert the parent repo '$sourceRepo' — $friendlyMsg"
                        $failCount++
                    }
                    $processedCount += $folderItems.Count
                    continue
                }
            }
        }

        foreach ($item in $folderItems) {
            $processedCount++
            $folderName = $item.Folder
            Write-Host ""
            Write-Host "  Working on $processedCount of $($migrateRows.Count)..." -ForegroundColor Cyan
            Write-Host "    Extracting folder: $sourceRepo / $folderName → new repo '$($item.NewRepoName)'" -ForegroundColor White

            try {
                $folderOutputName = $item.NewRepoName

                $mappings = @{
                    "`$/$sourceProject/$folderName" = $folderOutputName
                }

                & "$PSScriptRoot/Split-TfvcToGitRepos.ps1" `
                    -ConfigPath $ConfigPath `
                    -Collection $sourceCollection `
                    -ProjectName $sourceProject `
                    -TfvcPath "`$/$sourceProject" `
                    -FolderMappings $mappings

                Write-Host "    Moving to $TargetCollection / $TargetProject..." -ForegroundColor DarkGray

                & "$PSScriptRoot/Move-RepoToCollection.ps1" `
                    -ConfigPath $ConfigPath `
                    -SourceCollection $sourceCollection `
                    -SourceProject $sourceProject `
                    -TfvcPath "`$/$sourceProject/$folderName" `
                    -TargetCollection $TargetCollection `
                    -TargetProject $TargetProject `
                    -TargetRepoName $folderOutputName

                $item.Status = 'Success'
                $successCount++
                Write-Host "    ✓ Done" -ForegroundColor Green
                Write-MigrationLog -Message "Successfully extracted '$folderName' as '$folderOutputName'" -LogFile $logFile -Level SUCCESS
            }
            catch {
                $item.Status = 'Failed'
                $item.Error = Get-FriendlyError -ErrorMessage $_.Exception.Message
                $failCount++
                Write-Host "    ✗ Failed: $($item.Error)" -ForegroundColor Red
                Write-MigrationLog -Message "Failed to extract '$folderName': $($_.Exception.Message)" -LogFile $logFile -Level ERROR
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
    startTime     = $startTime.ToString('o')
    endTime       = (Get-Date).ToString('o')
    duration      = $totalDuration.ToString('hh\:mm\:ss')
    excelFile     = $ExcelPath
    worksheet     = $WorksheetName
    target        = "$TargetCollection/$TargetProject"
    totalRows     = $excelData.Count
    migrateCount  = $migrateRows.Count
    archiveCount  = $archiveRows.Count
    skipCount     = $skipRows.Count
    successCount  = $successCount
    failCount     = $failCount
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
Write-Host "  Results:" -ForegroundColor White
Write-Host "    ✓ Successful:    $successCount" -ForegroundColor Green
if ($failCount -gt 0) {
    Write-Host "    ✗ Failed:        $failCount" -ForegroundColor Red
}
Write-Host "    ○ Archived:      $($archiveRows.Count) (not migrated)" -ForegroundColor Yellow
Write-Host "    ○ Skipped:       $($skipRows.Count)" -ForegroundColor DarkGray
Write-Host ""

if ($failCount -gt 0) {
    Write-Host "  ┌─ Items that failed ───────────────────────────────────┐" -ForegroundColor Red
    foreach ($f in ($migrateRows | Where-Object { $_.Status -eq 'Failed' }) | Select-Object -First 15) {
        $label = "$($f.Collection)/$($f.Project)/$($f.Repo)"
        if ($f.Folder) { $label += "/$($f.Folder)" }
        Write-Host "  │  ✗ $label" -ForegroundColor Red
        Write-Host "  │    $($f.Error)" -ForegroundColor DarkGray
    }
    if ($failCount -gt 15) {
        Write-Host "  │  ... and $($failCount - 15) more (see manifest for full list)" -ForegroundColor Red
    }
    Write-Host "  └───────────────────────────────────────────────────────┘" -ForegroundColor Red
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
