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

.PARAMETER DryRun
    Preview what would be done without executing.

.PARAMETER Interactive
    Launch interactive mode — confirm settings via prompts.

.EXAMPLE
    ./Invoke-ExcelMigration.ps1 -ConfigPath ./config/migration-config.json `
        -ExcelPath ./excel-docs/MDR-4ADO-AllProjects.xlsx `
        -TargetCollection "ModernApps" -TargetProject "Platform" -DryRun

.EXAMPLE
    # Interactive mode
    ./Invoke-ExcelMigration.ps1 -ConfigPath ./config/migration-config.json -Interactive
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory)]
    [string]$ConfigPath,

    [string]$ExcelPath,

    [string]$WorksheetName = 'GAMS-Repos-App-Folder level',

    [string]$TargetCollection,

    [string]$TargetProject,

    [switch]$DryRun,

    [switch]$Interactive
)

$ErrorActionPreference = 'Stop'
Import-Module "$PSScriptRoot/modules/AdoTfvcMigrator.psm1" -Force

# ─── Check for ImportExcel Module ──────────────────────────────────────────────

if (-not (Get-Module -ListAvailable -Name ImportExcel)) {
    Write-Host ""
    Write-Host "  The 'ImportExcel' PowerShell module is required to read Excel files." -ForegroundColor Red
    Write-Host "  Install it with:  Install-Module ImportExcel -Scope CurrentUser -Force" -ForegroundColor Yellow
    Write-Host ""
    return
}
Import-Module ImportExcel

# ─── Load Config ───────────────────────────────────────────────────────────────

$config = Read-MigrationConfig -ConfigPath $ConfigPath
$logFile = Initialize-MigrationLog -LogDirectory $config.logDirectory -ScriptName 'ExcelMigration'

# ─── Interactive Mode ──────────────────────────────────────────────────────────

if ($Interactive) {
    Show-MenuHeader -Title "Excel-Driven Batch Migration"
    Write-Host "  This tool reads the MDR-4ADO-AllProjects Excel file and processes" -ForegroundColor DarkGray
    Write-Host "  each repository according to the Recommendation column." -ForegroundColor DarkGray
    Write-Host ""

    # Excel file path
    if (-not $ExcelPath) {
        $defaultExcel = Join-Path $PSScriptRoot 'excel-docs/MDR-4ADO-AllProjects.xlsx'
        if (Test-Path $defaultExcel) {
            Write-Host "  Found Excel file: $defaultExcel" -ForegroundColor Green
            Write-Host "  Use this file? [Y/n]: " -ForegroundColor Yellow -NoNewline
            $useDefault = Read-Host
            if ($useDefault.Trim() -match '^[Nn]') {
                Write-Host "  Enter path to MDR-4ADO-AllProjects.xlsx: " -ForegroundColor Yellow -NoNewline
                $ExcelPath = Read-Host
            }
            else {
                $ExcelPath = $defaultExcel
            }
        }
        else {
            Write-Host "  Enter path to MDR-4ADO-AllProjects.xlsx: " -ForegroundColor Yellow -NoNewline
            $ExcelPath = Read-Host
        }
    }

    # Target collection
    if (-not $TargetCollection) {
        Write-Host ""
        Write-Host "  Select the TARGET collection (where migrated repos will go):" -ForegroundColor White
        $TargetCollection = Select-AdoCollection -Config $config -Prompt 'Target collection'
        if (-not $TargetCollection) {
            Write-Host "  Cancelled." -ForegroundColor Yellow
            return
        }
    }

    # Target project
    if (-not $TargetProject) {
        Write-Host ""
        Write-Host "  Enter the target project name: " -ForegroundColor Yellow -NoNewline
        $TargetProject = Read-Host
        if (-not $TargetProject) {
            Write-Host "  Cancelled." -ForegroundColor Yellow
            return
        }
    }

    # Dry run?
    Write-Host ""
    Write-Host "  Run as dry-run first (preview only, no changes)? [Y/n]: " -ForegroundColor Yellow -NoNewline
    $dryChoice = Read-Host
    if (-not ($dryChoice.Trim() -match '^[Nn]')) {
        $DryRun = $true
    }
}

# ─── Validate Inputs ──────────────────────────────────────────────────────────

if (-not $ExcelPath -or -not (Test-Path $ExcelPath)) {
    throw "Excel file not found: $ExcelPath"
}
if (-not $TargetCollection) {
    throw "TargetCollection is required."
}
if (-not $TargetProject) {
    throw "TargetProject is required."
}

# ─── Read Excel Data ──────────────────────────────────────────────────────────

Write-MigrationLog -Message "Reading Excel file: $ExcelPath (sheet: $WorksheetName)" -LogFile $logFile -Level INFO

$excelData = Import-Excel -Path $ExcelPath -WorksheetName $WorksheetName

if (-not $excelData -or $excelData.Count -eq 0) {
    throw "No data found in worksheet '$WorksheetName'."
}

Write-MigrationLog -Message "Loaded $($excelData.Count) rows from Excel" -LogFile $logFile -Level INFO

# ─── Classify Rows ─────────────────────────────────────────────────────────────

$migrateRows = [System.Collections.ArrayList]::new()
$archiveRows = [System.Collections.ArrayList]::new()
$skipRows = [System.Collections.ArrayList]::new()

foreach ($row in $excelData) {
    $recommendation = ($row.Recommendation ?? '').ToString().Trim()
    $repoOrFolder = ($row.'Repo or Floder' ?? '').ToString().Trim()
    $reposType = ($row.'Repos Type' ?? '').ToString().Trim()
    $collection = ($row.Collection ?? '').ToString().Trim()
    $project = ($row.'Projects ' ?? $row.Projects ?? '').ToString().Trim()
    $repoName = ($row.Repos ?? '').ToString().Trim()
    $folderName = ($row.'Applications/Folders' ?? '').ToString().Trim()

    # Skip rows with no collection/project data
    if (-not $collection -or -not $project) {
        [void]$skipRows.Add([PSCustomObject]@{
            RowNo        = $row.No
            Collection   = $collection
            Project      = $project
            Repo         = $repoName
            Folder       = $folderName
            RepoType     = $reposType
            RepoOrFolder = $repoOrFolder
            Recommendation = $recommendation
            Action       = 'Skip'
            Reason       = 'Missing collection or project'
            Status       = 'Skipped'
            Error        = ''
        })
        continue
    }

    # Exclude BuildProcessTemplates
    if ($folderName -eq 'BuildProcessTemplates') {
        [void]$skipRows.Add([PSCustomObject]@{
            RowNo        = $row.No
            Collection   = $collection
            Project      = $project
            Repo         = $repoName
            Folder       = $folderName
            RepoType     = $reposType
            RepoOrFolder = $repoOrFolder
            Recommendation = $recommendation
            Action       = 'Skip'
            Reason       = 'BuildProcessTemplates excluded'
            Status       = 'Skipped'
            Error        = ''
        })
        continue
    }

    # Check recommendation
    if ($recommendation -match '^Archive') {
        [void]$archiveRows.Add([PSCustomObject]@{
            RowNo        = $row.No
            Collection   = $collection
            Project      = $project
            Repo         = $repoName
            Folder       = $folderName
            RepoType     = $reposType
            RepoOrFolder = $repoOrFolder
            Recommendation = $recommendation
            Action       = 'Archive'
            Reason       = "Recommendation: $recommendation"
            Status       = 'Skipped'
            Error        = ''
        })
    }
    elseif ($recommendation -match '(?i)^migrate') {
        $action = if ($repoOrFolder -eq 'Folder') { 'SpinOutFolder' } else { 'MigrateRepo' }
        [void]$migrateRows.Add([PSCustomObject]@{
            RowNo        = $row.No
            Collection   = $collection
            Project      = $project
            Repo         = $repoName
            Folder       = $folderName
            RepoType     = $reposType
            RepoOrFolder = $repoOrFolder
            Recommendation = $recommendation
            Action       = $action
            Reason       = "Recommendation: $recommendation / Type: $repoOrFolder"
            Status       = 'Pending'
            Error        = ''
        })
    }
    else {
        [void]$skipRows.Add([PSCustomObject]@{
            RowNo        = $row.No
            Collection   = $collection
            Project      = $project
            Repo         = $repoName
            Folder       = $folderName
            RepoType     = $reposType
            RepoOrFolder = $repoOrFolder
            Recommendation = $recommendation
            Action       = 'Skip'
            Reason       = "Unrecognized recommendation: $recommendation"
            Status       = 'Skipped'
            Error        = ''
        })
    }
}

# ─── Summary ───────────────────────────────────────────────────────────────────

Write-Host ""
Write-Host "═══════════════════════════════════════════════════════" -ForegroundColor Cyan
Write-Host "  Excel Migration Plan Summary" -ForegroundColor Cyan
Write-Host "═══════════════════════════════════════════════════════" -ForegroundColor Cyan
Write-Host ""
Write-Host "  Total rows:          $($excelData.Count)" -ForegroundColor White
Write-Host "  To migrate:          $($migrateRows.Count)" -ForegroundColor Green
$repoMigrations = @($migrateRows | Where-Object { $_.Action -eq 'MigrateRepo' })
$folderSpinouts = @($migrateRows | Where-Object { $_.Action -eq 'SpinOutFolder' })
Write-Host "    Full repos:        $($repoMigrations.Count)" -ForegroundColor DarkGray
Write-Host "    Folder spin-outs:  $($folderSpinouts.Count)" -ForegroundColor DarkGray
Write-Host "  Archive (skip):      $($archiveRows.Count)" -ForegroundColor Yellow
Write-Host "  Other skips:         $($skipRows.Count)" -ForegroundColor DarkGray
Write-Host ""
Write-Host "  Target: $TargetCollection / $TargetProject" -ForegroundColor White
Write-Host ""

if ($DryRun) {
    Write-Host "═══════════════════════ DRY RUN ═══════════════════════" -ForegroundColor Yellow
    Write-Host ""

    if ($repoMigrations.Count -gt 0) {
        Write-Host "  Repos to migrate:" -ForegroundColor White
        foreach ($r in $repoMigrations) {
            $convertNote = if ($r.RepoType -eq 'TFVC') { ' (TFVC→Git)' } else { '' }
            Write-Host "    $($r.Collection)/$($r.Project)/$($r.Repo)$convertNote → $TargetCollection/$TargetProject" -ForegroundColor DarkGray
        }
        Write-Host ""
    }

    if ($folderSpinouts.Count -gt 0) {
        Write-Host "  Folders to spin out:" -ForegroundColor White
        foreach ($r in $folderSpinouts) {
            Write-Host "    $($r.Collection)/$($r.Project)/$($r.Repo)/$($r.Folder) → new repo '$($r.Folder)' in $TargetCollection/$TargetProject" -ForegroundColor DarkGray
        }
        Write-Host ""
    }

    if ($archiveRows.Count -gt 0) {
        Write-Host "  Repos marked for archive (will be skipped):" -ForegroundColor Yellow
        foreach ($r in $archiveRows | Select-Object -First 20) {
            Write-Host "    $($r.Collection)/$($r.Project)/$($r.Repo) — $($r.Recommendation)" -ForegroundColor DarkGray
        }
        if ($archiveRows.Count -gt 20) {
            Write-Host "    ... and $($archiveRows.Count - 20) more" -ForegroundColor DarkGray
        }
        Write-Host ""
    }

    Write-Host "═══════════════════ END DRY RUN ═══════════════════════" -ForegroundColor Yellow
    Write-Host "  Re-run without -DryRun to execute." -ForegroundColor Cyan
    Write-Host ""

    # Still produce the manifest for dry run
    $allActions = [System.Collections.ArrayList]::new()
    foreach ($r in $migrateRows) { [void]$allActions.Add($r) }
    foreach ($r in $archiveRows) { [void]$allActions.Add($r) }
    foreach ($r in $skipRows) { [void]$allActions.Add($r) }

    $manifestPath = Join-Path $config.outputDirectory "excel-migration-manifest-DRYRUN-$(Get-Date -Format 'yyyyMMdd-HHmmss').csv"
    $allActions | Export-Csv -Path $manifestPath -NoTypeInformation -Encoding utf8
    Write-MigrationLog -Message "Dry-run manifest written to: $manifestPath" -LogFile $logFile -Level SUCCESS
    Write-Host "  Manifest: $manifestPath" -ForegroundColor White
    return
}

# ─── Execute Migrations ───────────────────────────────────────────────────────

Write-Host "  Starting migration of $($migrateRows.Count) items..." -ForegroundColor Cyan
Write-Host ""

$startTime = Get-Date
$processedCount = 0

# Group migrations by Collection/Project/Repo to batch operations on same repo
$grouped = $migrateRows | Group-Object -Property { "$($_.Collection)|$($_.Project)|$($_.Repo)" }

foreach ($group in $grouped) {
    $parts = $group.Name -split '\|'
    $sourceCollection = $parts[0]
    $sourceProject = $parts[1]
    $sourceRepo = $parts[2]

    $pat = $config.collections[$sourceCollection]?.pat
    if (-not $pat) {
        Write-MigrationLog -Message "Collection '$sourceCollection' not found in config — skipping group" -LogFile $logFile -Level WARN
        foreach ($item in $group.Group) {
            $item.Status = 'Failed'
            $item.Error = "Collection '$sourceCollection' not in config"
        }
        continue
    }

    # Separate full-repo migrations from folder spin-outs in this group
    $repoItems = @($group.Group | Where-Object { $_.Action -eq 'MigrateRepo' })
    $folderItems = @($group.Group | Where-Object { $_.Action -eq 'SpinOutFolder' })

    # ── Full Repo Migration ──
    foreach ($item in $repoItems) {
        $processedCount++
        Write-Host "  [$processedCount/$($migrateRows.Count)] Migrating repo: $sourceCollection/$sourceProject/$sourceRepo" -ForegroundColor White

        try {
            $needsConversion = $item.RepoType -eq 'TFVC'
            $outputName = $sourceRepo -replace '[^a-zA-Z0-9_-]', '-'

            if ($needsConversion) {
                Write-MigrationLog -Message "Converting TFVC repo '$sourceRepo' to Git" -LogFile $logFile -Level INFO

                & "$PSScriptRoot/Convert-TfvcToGit.ps1" `
                    -ConfigPath $ConfigPath `
                    -Collection $sourceCollection `
                    -ProjectName $sourceProject `
                    -TfvcPath "`$/$sourceProject" `
                    -OutputRepoName $outputName
            }

            # Move to target collection
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
            Write-MigrationLog -Message "Successfully migrated '$sourceRepo'" -LogFile $logFile -Level SUCCESS
        }
        catch {
            $item.Status = 'Failed'
            $item.Error = $_.Exception.Message
            Write-MigrationLog -Message "Failed to migrate '$sourceRepo': $($_.Exception.Message)" -LogFile $logFile -Level ERROR
        }
    }

    # ── Folder Spin-Outs ──
    if ($folderItems.Count -gt 0) {
        # First convert the parent repo if it's TFVC
        $parentConverted = $false
        $parentRepoType = $folderItems[0].RepoType

        if ($parentRepoType -eq 'TFVC') {
            try {
                $parentOutputName = $sourceRepo -replace '[^a-zA-Z0-9_-]', '-'

                # Check if we already converted this repo in the full-repo step
                $parentRepoPath = Join-Path $config.outputDirectory $parentOutputName
                if (-not (Test-Path $parentRepoPath)) {
                    Write-MigrationLog -Message "Converting parent TFVC repo '$sourceRepo' to Git for folder extraction" -LogFile $logFile -Level INFO

                    & "$PSScriptRoot/Convert-TfvcToGit.ps1" `
                        -ConfigPath $ConfigPath `
                        -Collection $sourceCollection `
                        -ProjectName $sourceProject `
                        -TfvcPath "`$/$sourceProject" `
                        -OutputRepoName $parentOutputName
                }
                $parentConverted = $true
            }
            catch {
                Write-MigrationLog -Message "Failed to convert parent repo '$sourceRepo': $($_.Exception.Message)" -LogFile $logFile -Level ERROR
                foreach ($fi in $folderItems) {
                    $fi.Status = 'Failed'
                    $fi.Error = "Parent repo conversion failed: $($_.Exception.Message)"
                }
                continue
            }
        }

        foreach ($item in $folderItems) {
            $processedCount++
            $folderName = $item.Folder
            Write-Host "  [$processedCount/$($migrateRows.Count)] Spinning out folder: $sourceCollection/$sourceProject/$sourceRepo/$folderName" -ForegroundColor White

            try {
                $folderOutputName = $folderName -replace '[^a-zA-Z0-9_-]', '-'

                # Use Split-TfvcToGitRepos to extract the folder
                $mappings = @{
                    "`$/$sourceProject/$folderName" = $folderOutputName
                }

                & "$PSScriptRoot/Split-TfvcToGitRepos.ps1" `
                    -ConfigPath $ConfigPath `
                    -Collection $sourceCollection `
                    -ProjectName $sourceProject `
                    -TfvcPath "`$/$sourceProject" `
                    -FolderMappings $mappings

                # Move to target collection
                Write-MigrationLog -Message "Moving spun-out repo '$folderOutputName' to $TargetCollection/$TargetProject" -LogFile $logFile -Level INFO

                & "$PSScriptRoot/Move-RepoToCollection.ps1" `
                    -ConfigPath $ConfigPath `
                    -SourceCollection $sourceCollection `
                    -SourceProject $sourceProject `
                    -TfvcPath "`$/$sourceProject/$folderName" `
                    -TargetCollection $TargetCollection `
                    -TargetProject $TargetProject `
                    -TargetRepoName $folderOutputName

                $item.Status = 'Success'
                Write-MigrationLog -Message "Successfully spun out '$folderName' as '$folderOutputName'" -LogFile $logFile -Level SUCCESS
            }
            catch {
                $item.Status = 'Failed'
                $item.Error = $_.Exception.Message
                Write-MigrationLog -Message "Failed to spin out '$folderName': $($_.Exception.Message)" -LogFile $logFile -Level ERROR
            }
        }
    }
}

# ─── Comprehensive Manifest ──────────────────────────────────────────────────

$totalDuration = (Get-Date) - $startTime

# Combine all rows into one manifest
$allActions = [System.Collections.ArrayList]::new()
foreach ($r in $migrateRows) {
    $r | Add-Member -NotePropertyName Destination -NotePropertyValue "$TargetCollection/$TargetProject" -Force
    $r | Add-Member -NotePropertyName ConvertedToGit -NotePropertyValue ($r.RepoType -eq 'TFVC') -Force
    $r | Add-Member -NotePropertyName SpunOut -NotePropertyValue ($r.Action -eq 'SpinOutFolder') -Force
    [void]$allActions.Add($r)
}
foreach ($r in $archiveRows) {
    $r | Add-Member -NotePropertyName Destination -NotePropertyValue 'N/A (Archived)' -Force
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

# Output summary
$successCount = @($migrateRows | Where-Object { $_.Status -eq 'Success' }).Count
$failCount = @($migrateRows | Where-Object { $_.Status -eq 'Failed' }).Count
$pendingCount = @($migrateRows | Where-Object { $_.Status -eq 'Pending' }).Count

Write-Host ""
Write-Host "═══════════════════════════════════════════════════════" -ForegroundColor Cyan
Write-Host "  Migration Complete" -ForegroundColor Cyan
Write-Host "═══════════════════════════════════════════════════════" -ForegroundColor Cyan
Write-Host ""
Write-Host "  Duration:        $($totalDuration.ToString('hh\:mm\:ss'))" -ForegroundColor White
Write-Host "  Total processed: $($migrateRows.Count)" -ForegroundColor White
Write-Host "  Successful:      $successCount" -ForegroundColor Green
if ($failCount -gt 0) {
    Write-Host "  Failed:          $failCount" -ForegroundColor Red
}
Write-Host "  Archived (skip): $($archiveRows.Count)" -ForegroundColor Yellow
Write-Host "  Other skips:     $($skipRows.Count)" -ForegroundColor DarkGray
Write-Host ""

# Show failures
if ($failCount -gt 0) {
    Write-Host "  Failed items:" -ForegroundColor Red
    foreach ($f in ($migrateRows | Where-Object { $_.Status -eq 'Failed' })) {
        Write-Host "    $($f.Collection)/$($f.Project)/$($f.Repo)$(if($f.Folder){"/$($f.Folder)"}): $($f.Error)" -ForegroundColor Red
    }
    Write-Host ""
}

# Write CSV manifest
$manifestPath = Join-Path $config.outputDirectory "excel-migration-manifest-$(Get-Date -Format 'yyyyMMdd-HHmmss').csv"
$allActions | Select-Object RowNo, Collection, Project, Repo, Folder, RepoType, RepoOrFolder, `
    Recommendation, Action, Reason, Status, Error, Destination, ConvertedToGit, SpunOut |
    Export-Csv -Path $manifestPath -NoTypeInformation -Encoding utf8

Write-MigrationLog -Message "Migration manifest written to: $manifestPath" -LogFile $logFile -Level SUCCESS
Write-Host "  Manifest: $manifestPath" -ForegroundColor White

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
Write-MigrationLog -Message "Report saved: $reportPath" -LogFile $logFile -Level INFO

Write-Host "  Report:   $reportPath" -ForegroundColor White
Write-Host "  Log:      $logFile" -ForegroundColor DarkGray
Write-Host ""
