#Requires -Version 7.0
<#
.SYNOPSIS
    Orchestrates a batch migration: discover → convert → push to GitHub.

.DESCRIPTION
    Reads a migration plan (JSON) that defines which TFVC paths to migrate,
    how to split them, and where to push them. Executes the plan end-to-end
    with logging and a summary report.

.PARAMETER ConfigPath
    Path to migration-config.json.

.PARAMETER PlanFile
    Path to a migration plan JSON file.

.PARAMETER DryRun
    Validate the plan and show what would be done without executing.

.PARAMETER StopOnError
    Stop processing on the first error. Default: continue and report failures.

.NOTES
    Migration Plan Format (migration-plan.json):

    {
        "migrations": [
            {
                "name": "Legacy App Full Repo",
                "type": "full",
                "sourceCollection": "GAMS",
                "sourceProject": "LegacyApp",
                "tfvcPath": "$/LegacyApp",
                "gitHubOrg": "Contoso",
                "gitHubRepo": "legacy-app",
                "description": "Full legacy app migration"
            },
            {
                "name": "MonoRepo Split",
                "type": "split",
                "sourceCollection": "GAMS",
                "sourceProject": "MonoRepo",
                "tfvcPath": "$/MonoRepo",
                "splits": [
                    {
                        "tfvcFolder": "$/MonoRepo/ServiceA",
                        "gitHubOrg": "Contoso",
                        "gitHubRepo": "service-a"
                    },
                    {
                        "tfvcFolder": "$/MonoRepo/ServiceB",
                        "gitHubOrg": "Contoso",
                        "gitHubRepo": "service-b"
                    }
                ]
            },
            {
                "name": "Cross-Collection Move",
                "type": "move",
                "sourceCollection": "GAMS",
                "sourceProject": "OldProject",
                "tfvcPath": "$/OldProject/Component",
                "targetCollection": "ModernApps",
                "targetProject": "Platform",
                "targetRepoName": "component",
                "gitHubOrg": "Contoso",
                "gitHubRepo": "component"
            }
        ]
    }

.EXAMPLE
    ./Start-Migration.ps1 -ConfigPath ./config/migration-config.json `
        -PlanFile ./config/migration-plan.json -DryRun

.EXAMPLE
    ./Start-Migration.ps1 -ConfigPath ./config/migration-config.json `
        -PlanFile ./config/migration-plan.json
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory)]
    [string]$ConfigPath,

    [Parameter(Mandatory)]
    [string]$PlanFile,

    [switch]$DryRun,

    [switch]$StopOnError
)

$ErrorActionPreference = 'Stop'
Import-Module "$PSScriptRoot/modules/AdoTfvcMigrator.psm1" -Force

# ─── Load Config & Plan ───────────────────────────────────────────────────────

$config = Read-MigrationConfig -ConfigPath $ConfigPath
$logFile = Initialize-MigrationLog -LogDirectory $config.logDirectory -ScriptName 'StartMigration'

if (-not (Test-Path $PlanFile)) {
    throw "Migration plan not found: $PlanFile"
}

$plan = Get-Content $PlanFile -Raw | ConvertFrom-Json -AsHashtable

if (-not $plan.migrations -or $plan.migrations.Count -eq 0) {
    throw "Migration plan has no entries."
}

$totalCount = $plan.migrations.Count
Write-MigrationLog -Message "Migration plan loaded: $totalCount entries" -LogFile $logFile -Level INFO

if ($DryRun) {
    Write-Host ""
    Write-Host "═══════════════════════ DRY RUN ═══════════════════════" -ForegroundColor Yellow
    Write-Host ""
}

# ─── Validate Plan ────────────────────────────────────────────────────────────

Write-MigrationLog -Message "Validating migration plan..." -LogFile $logFile -Level INFO

$validationErrors = [System.Collections.ArrayList]::new()

foreach ($migration in $plan.migrations) {
    $name = $migration.name

    if (-not $migration.type) {
        [void]$validationErrors.Add("[$name] Missing 'type' (full|split|move)")
    }
    if (-not $migration.sourceCollection) {
        [void]$validationErrors.Add("[$name] Missing 'sourceCollection'")
    }
    elseif (-not $config.collections[$migration.sourceCollection]) {
        [void]$validationErrors.Add("[$name] Collection '$($migration.sourceCollection)' not in config")
    }
    if (-not $migration.tfvcPath) {
        [void]$validationErrors.Add("[$name] Missing 'tfvcPath'")
    }

    switch ($migration.type) {
        'split' {
            if (-not $migration.splits -or $migration.splits.Count -eq 0) {
                [void]$validationErrors.Add("[$name] Split migration has no 'splits' entries")
            }
        }
        'move' {
            if (-not $migration.targetCollection) {
                [void]$validationErrors.Add("[$name] Move migration missing 'targetCollection'")
            }
            if (-not $migration.targetProject) {
                [void]$validationErrors.Add("[$name] Move migration missing 'targetProject'")
            }
        }
    }
}

if ($validationErrors.Count -gt 0) {
    Write-MigrationLog -Message "Validation errors:" -LogFile $logFile -Level ERROR
    foreach ($err in $validationErrors) {
        Write-MigrationLog -Message "  $err" -LogFile $logFile -Level ERROR
    }
    throw "Migration plan has $($validationErrors.Count) validation error(s). Fix them and retry."
}

Write-MigrationLog -Message "Validation passed" -LogFile $logFile -Level SUCCESS

# ─── Show Plan Summary ────────────────────────────────────────────────────────

Write-Host ""
Write-Host "Migration Plan:" -ForegroundColor Cyan
Write-Host "─────────────────────────────────────────────────────" -ForegroundColor DarkGray

foreach ($i in 0..($plan.migrations.Count - 1)) {
    $m = $plan.migrations[$i]
    $idx = $i + 1
    Write-Host "  [$idx/$totalCount] $($m.name)" -ForegroundColor White
    Write-Host "          Type:   $($m.type)" -ForegroundColor DarkGray
    Write-Host "          Source: $($m.sourceCollection)/$($m.sourceProject) ($($m.tfvcPath))" -ForegroundColor DarkGray

    switch ($m.type) {
        'full' {
            Write-Host "          Target: GitHub $($m.gitHubOrg)/$($m.gitHubRepo)" -ForegroundColor DarkGray
        }
        'split' {
            foreach ($s in $m.splits) {
                Write-Host "          Split:  $($s.tfvcFolder) → GitHub $($s.gitHubOrg)/$($s.gitHubRepo)" -ForegroundColor DarkGray
            }
        }
        'move' {
            Write-Host "          Move:   → $($m.targetCollection)/$($m.targetProject)/$($m.targetRepoName)" -ForegroundColor DarkGray
            if ($m.gitHubRepo) {
                Write-Host "          Then:   → GitHub $($m.gitHubOrg)/$($m.gitHubRepo)" -ForegroundColor DarkGray
            }
        }
    }
    Write-Host ""
}

if ($DryRun) {
    Write-Host "═══════════════════ END DRY RUN ═══════════════════════" -ForegroundColor Yellow
    Write-Host "Re-run without -DryRun to execute." -ForegroundColor Cyan
    return
}

# ─── Execute Plan ──────────────────────────────────────────────────────────────

$results = [System.Collections.ArrayList]::new()
$startTime = Get-Date

foreach ($i in 0..($plan.migrations.Count - 1)) {
    $m = $plan.migrations[$i]
    $idx = $i + 1
    $name = $m.name

    Write-Host ""
    Write-Host "═══════════════════════════════════════════════════════" -ForegroundColor Cyan
    Write-MigrationLog -Message "[$idx/$totalCount] Starting: $name" -LogFile $logFile -Level INFO
    Write-Host "═══════════════════════════════════════════════════════" -ForegroundColor Cyan

    $migrationStart = Get-Date
    $status = 'Success'
    $errorMsg = ''

    try {
        switch ($m.type) {
            'full' {
                # Convert TFVC to Git
                & "$PSScriptRoot/Convert-TfvcToGit.ps1" `
                    -ConfigPath $ConfigPath `
                    -Collection $m.sourceCollection `
                    -ProjectName $m.sourceProject `
                    -TfvcPath $m.tfvcPath `
                    -OutputRepoName $m.gitHubRepo

                # Push to GitHub
                if ($m.gitHubOrg -and $m.gitHubRepo) {
                    $repoPath = Join-Path $config.outputDirectory $m.gitHubRepo
                    & "$PSScriptRoot/Push-ToGitHub.ps1" `
                        -ConfigPath $ConfigPath `
                        -RepoPath $repoPath `
                        -GitHubOrg $m.gitHubOrg `
                        -GitHubRepo $m.gitHubRepo `
                        -Description ($m.description ?? "Migrated from $($m.sourceCollection)/$($m.tfvcPath)")
                }
            }

            'split' {
                # Build folder mappings
                $mappings = @{}
                foreach ($s in $m.splits) {
                    $mappings[$s.tfvcFolder] = $s.gitHubRepo
                }

                & "$PSScriptRoot/Split-TfvcToGitRepos.ps1" `
                    -ConfigPath $ConfigPath `
                    -Collection $m.sourceCollection `
                    -ProjectName $m.sourceProject `
                    -TfvcPath $m.tfvcPath `
                    -FolderMappings $mappings

                # Push each split repo to GitHub
                foreach ($s in $m.splits) {
                    if ($s.gitHubOrg -and $s.gitHubRepo) {
                        $repoPath = Join-Path $config.outputDirectory $s.gitHubRepo
                        & "$PSScriptRoot/Push-ToGitHub.ps1" `
                            -ConfigPath $ConfigPath `
                            -RepoPath $repoPath `
                            -GitHubOrg $s.gitHubOrg `
                            -GitHubRepo $s.gitHubRepo `
                            -Description "Migrated from $($m.sourceCollection)/$($s.tfvcFolder)"
                    }
                }
            }

            'move' {
                # Move across collections
                & "$PSScriptRoot/Move-RepoToCollection.ps1" `
                    -ConfigPath $ConfigPath `
                    -SourceCollection $m.sourceCollection `
                    -SourceProject $m.sourceProject `
                    -TfvcPath $m.tfvcPath `
                    -TargetCollection $m.targetCollection `
                    -TargetProject $m.targetProject `
                    -TargetRepoName $m.targetRepoName

                # Optionally push to GitHub as well
                if ($m.gitHubOrg -and $m.gitHubRepo) {
                    $repoPath = Join-Path $config.outputDirectory $m.targetRepoName
                    & "$PSScriptRoot/Push-ToGitHub.ps1" `
                        -ConfigPath $ConfigPath `
                        -RepoPath $repoPath `
                        -GitHubOrg $m.gitHubOrg `
                        -GitHubRepo $m.gitHubRepo `
                        -Description ($m.description ?? "Migrated from $($m.sourceCollection)/$($m.tfvcPath)")
                }
            }

            default {
                throw "Unknown migration type: $($m.type)"
            }
        }
    }
    catch {
        $errorMsg = $_.Exception.Message

        # Detect path-too-long errors specifically
        if ($errorMsg -match 'too long|path.*long|248 char|260 char|TF400959|PathTooLong|could not find a part of the path') {
            $status = 'PathTooLong'
            $errorMsg = "File paths exceed the Windows 260-character limit. " +
                "Try migrating specific subfolders instead of the entire repo, " +
                "or rename deeply nested paths in TFVC first."
        }
        else {
            $status = 'Failed'
        }

        Write-MigrationLog -Message "[$idx/$totalCount] $($status): $name — $errorMsg" -LogFile $logFile -Level ERROR

        if ($StopOnError) {
            throw "Migration stopped on error at [$idx/$totalCount] $name"
        }
    }

    $duration = (Get-Date) - $migrationStart

    [void]$results.Add([PSCustomObject]@{
        Index    = $idx
        Name     = $name
        Type     = $m.type
        Status   = $status
        Duration = $duration.ToString('hh\:mm\:ss')
        Error    = $errorMsg
    })

    Write-MigrationLog -Message "[$idx/$totalCount] $status — $name ($($duration.ToString('hh\:mm\:ss')))" -LogFile $logFile -Level $(if ($status -eq 'Success') { 'SUCCESS' } else { 'ERROR' })
}

# ─── Final Report ──────────────────────────────────────────────────────────────

$totalDuration = (Get-Date) - $startTime
$successCount = ($results | Where-Object { $_.Status -eq 'Success' }).Count
$failCount = ($results | Where-Object { $_.Status -eq 'Failed' }).Count

Write-Host ""
Write-Host "═══════════════════════════════════════════════════════" -ForegroundColor Cyan
Write-Host "  Migration Complete" -ForegroundColor Cyan
Write-Host "═══════════════════════════════════════════════════════" -ForegroundColor Cyan
Write-Host ""

$results | Format-Table Index, Name, Type, Status, Duration -AutoSize

Write-Host "Total:    $totalCount migrations" -ForegroundColor White
Write-Host "Success:  $successCount" -ForegroundColor Green
if ($failCount -gt 0) {
    Write-Host "Failed:   $failCount" -ForegroundColor Red
    Write-Host ""
    Write-Host "Failed migrations:" -ForegroundColor Red
    $results | Where-Object { $_.Status -eq 'Failed' } | ForEach-Object {
        Write-Host "  [$($_.Index)] $($_.Name): $($_.Error)" -ForegroundColor Red
    }
}
Write-Host "Duration: $($totalDuration.ToString('hh\:mm\:ss'))" -ForegroundColor White
Write-Host "Log:      $logFile" -ForegroundColor DarkGray

# Write report to file
$reportPath = Join-Path $config.outputDirectory "migration-report-$(Get-Date -Format 'yyyyMMdd-HHmmss').json"
$report = @{
    startTime    = $startTime.ToString('o')
    endTime      = (Get-Date).ToString('o')
    duration     = $totalDuration.ToString('hh\:mm\:ss')
    totalCount   = $totalCount
    successCount = $successCount
    failCount    = $failCount
    results      = $results
}
$report | ConvertTo-Json -Depth 5 | Out-File $reportPath -Encoding utf8
Write-MigrationLog -Message "Report saved: $reportPath" -LogFile $logFile -Level INFO
