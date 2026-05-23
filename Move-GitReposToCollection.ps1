#Requires -Version 7.0
<#
.SYNOPSIS
    Bulk-moves all Git repositories from a source ADO collection/project into a target
    project in a DIFFERENT ADO collection, preserving all branches, tags, and history.

.DESCRIPTION
    For each Git repo in the source project this script performs a mirror clone, creates
    a matching repo in the target collection/project (if it doesn't already exist), then
    pushes the mirror to the target. Branches, tags, and notes are all preserved.

    Use cases:
      • Consolidate Git repos from a legacy collection into a modern one.
      • Reorganize project ownership across ADO Server collections.
      • Stage a bulk move before pointing CI/CD at the new location.

    Supports:
      • -Interactive mode with numbered menus.
      • -DryRun mode (lists what would happen without making any changes).
      • Include/Exclude filters on repo names (wildcards supported).
      • Skip-existing toggle (don't touch repos that already exist in the target).
      • Continue-on-error with a manifest CSV summarizing each repo's outcome.

.PARAMETER ConfigPath
    Path to migration-config.json.

.PARAMETER Interactive
    Launch interactive mode — browse and select source/target via menus, optionally
    multi-select the repos to move.

.PARAMETER SourceCollection
    Source ADO collection name (must exist in config.collections).

.PARAMETER SourceProject
    Source team project name to enumerate Git repos from.

.PARAMETER TargetCollection
    Target ADO collection name (must exist in config.collections, must differ from source).

.PARAMETER TargetProject
    Target team project name. Must already exist in the target collection.

.PARAMETER IncludeRepoNames
    Optional array of repo names (or wildcard patterns, e.g. 'web-*') to include.
    If omitted, all repos in the source project are eligible.

.PARAMETER ExcludeRepoNames
    Optional array of repo names (or wildcard patterns) to exclude.

.PARAMETER SkipExisting
    If a repo with the same name already exists in the target, skip it instead of
    pushing into it.

.PARAMETER DryRun
    Print the move plan and exit without touching any repos.

.PARAMETER WorkingDirectory
    Override the directory used for the local mirror clones. Defaults to
    $config.outputDirectory/git-bulk-move/$SourceCollection-$SourceProject.

.PARAMETER KeepWorkingCopies
    Don't delete the local mirror clones after a successful push.

.EXAMPLE
    # Interactive mode — guided menus, with multi-select of repos
    ./Move-GitReposToCollection.ps1 -ConfigPath ./config/migration-config.json -Interactive

.EXAMPLE
    # Dry run — preview what would move from GAMS/LegacyApp to Modern/Platform
    ./Move-GitReposToCollection.ps1 -ConfigPath ./config/migration-config.json `
        -SourceCollection GAMS -SourceProject LegacyApp `
        -TargetCollection Modern -TargetProject Platform -DryRun

.EXAMPLE
    # Direct mode — move every repo, but skip any that already exist in the target
    ./Move-GitReposToCollection.ps1 -ConfigPath ./config/migration-config.json `
        -SourceCollection GAMS -SourceProject LegacyApp `
        -TargetCollection Modern -TargetProject Platform -SkipExisting

.EXAMPLE
    # Direct mode — move only the 'web-*' repos, excluding archived ones
    ./Move-GitReposToCollection.ps1 -ConfigPath ./config/migration-config.json `
        -SourceCollection GAMS -SourceProject LegacyApp `
        -TargetCollection Modern -TargetProject Platform `
        -IncludeRepoNames 'web-*' -ExcludeRepoNames '*-archive'
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory)]
    [string]$ConfigPath,

    [switch]$Interactive,

    [string]$SourceCollection,
    [string]$SourceProject,
    [string]$TargetCollection,
    [string]$TargetProject,

    [string[]]$IncludeRepoNames,
    [string[]]$ExcludeRepoNames,

    [switch]$SkipExisting,
    [switch]$DryRun,

    [string]$WorkingDirectory,

    [switch]$KeepWorkingCopies
)

$ErrorActionPreference = 'Stop'
Import-Module "$PSScriptRoot/modules/AdoTfvcMigrator.psm1" -Force

# ─── Load Config ───────────────────────────────────────────────────────────────

$config = Read-MigrationConfig -ConfigPath $ConfigPath
$logFile = Initialize-MigrationLog -LogDirectory $config.logDirectory -ScriptName 'MoveGitReposToCollection'

# ─── Helpers ───────────────────────────────────────────────────────────────────

function Test-RepoNameMatchesAny {
    param(
        [Parameter(Mandatory)][string]$Name,
        [string[]]$Patterns
    )
    if (-not $Patterns -or $Patterns.Count -eq 0) { return $false }
    foreach ($pattern in $Patterns) {
        if ($Name -like $pattern) { return $true }
    }
    return $false
}

function Format-Bytes {
    param([long]$Bytes)
    if ($Bytes -ge 1GB) { return "$([math]::Round($Bytes / 1GB, 2)) GB" }
    if ($Bytes -ge 1MB) { return "$([math]::Round($Bytes / 1MB, 1)) MB" }
    if ($Bytes -ge 1KB) { return "$([math]::Round($Bytes / 1KB, 1)) KB" }
    return "$Bytes B"
}

# ─── Interactive Mode ──────────────────────────────────────────────────────────

$selectedRepoNames = $null  # null means "all that match filters"

if ($Interactive) {
    Show-MenuHeader -Title "Bulk Move Git Repos Between Collections"
    Write-Host "This wizard moves every Git repo in a source project to a target project" -ForegroundColor DarkGray
    Write-Host "in a DIFFERENT ADO collection. Branches, tags, and history are preserved." -ForegroundColor DarkGray

    # ── 1. Pick source collection ──
    $SourceCollection = Select-AdoCollection -Config $config -Prompt 'Select SOURCE collection'
    if (-not $SourceCollection) { Write-Host "Cancelled." -ForegroundColor Yellow; return }

    # ── 2. Pick source project ──
    $SourceProject = Select-AdoProject -Config $config -Collection $SourceCollection -Prompt 'Select SOURCE project'
    if (-not $SourceProject) { Write-Host "Cancelled." -ForegroundColor Yellow; return }

    # ── 3. Pick target collection (must differ from source) ──
    Show-MenuHeader -Title "Select DESTINATION"
    Write-Host "  Source: $SourceCollection / $SourceProject" -ForegroundColor DarkGray
    Write-Host ""

    $TargetCollection = Select-AdoCollection -Config $config -Prompt 'Select TARGET collection'
    if (-not $TargetCollection) { Write-Host "Cancelled." -ForegroundColor Yellow; return }

    if ($TargetCollection -eq $SourceCollection) {
        Write-Host ""
        Write-Host "  Source and target collections must be different." -ForegroundColor Red
        return
    }

    # ── 4. Pick target project ──
    $TargetProject = Select-AdoProject -Config $config -Collection $TargetCollection -Prompt 'Select TARGET project'
    if (-not $TargetProject) { Write-Host "Cancelled." -ForegroundColor Yellow; return }

    # ── 5. Enumerate source repos so we can let the user multi-select ──
    Show-MenuHeader -Title "Discovering source repositories"
    $sourcePat = $config.collections[$SourceCollection].pat
    Write-Host "  Listing Git repos in $SourceCollection / $SourceProject..." -ForegroundColor DarkGray
    $sourceRepos = Get-AdoGitRepositories -ServerUrl $config.adoServerUrl `
        -Collection $SourceCollection -Pat $sourcePat -ProjectName $SourceProject

    if (-not $sourceRepos -or $sourceRepos.Count -eq 0) {
        Write-Host "  No Git repos found in $SourceCollection / $SourceProject." -ForegroundColor Yellow
        return
    }

    $sortedRepos = $sourceRepos | Sort-Object name
    $displayItems = $sortedRepos | ForEach-Object {
        $size = if ($_.size) { Format-Bytes -Bytes $_.size } else { '?' }
        $branch = if ($_.defaultBranch) { ($_.defaultBranch -replace '^refs/heads/', '') } else { '(none)' }
        "$($_.name)  [$branch, $size]"
    }

    Write-Host ""
    Write-Host "  Found $($sortedRepos.Count) repo(s). Choose which to move:" -ForegroundColor White
    Write-Host "    [A] All repos" -ForegroundColor White
    Write-Host "    [S] Select specific repos (multi-select)" -ForegroundColor White
    Write-Host "    [0] Cancel" -ForegroundColor DarkGray
    Write-Host ""
    Write-Host "  Choice: " -ForegroundColor Yellow -NoNewline
    $modeChoice = (Read-Host).Trim().ToUpper()

    switch ($modeChoice) {
        '0' { Write-Host "Cancelled." -ForegroundColor Yellow; return }
        'S' {
            Show-MenuHeader -Title "Select repos to move"
            $picked = Show-NumberedMenu -Items $displayItems -Prompt 'Pick repos' -MultiSelect -AllowBack
            if ($null -eq $picked -or $picked.Count -eq 0) {
                Write-Host "Cancelled." -ForegroundColor Yellow; return
            }
            $selectedRepoNames = @($picked | ForEach-Object { $sortedRepos[$_].name })
        }
        default {
            # Treat anything else (A, empty, etc.) as "all"
            $selectedRepoNames = @($sortedRepos.name)
        }
    }

    # ── 6. Skip existing? ──
    Write-Host ""
    Write-Host "  If a repo already exists in the target, skip it? [y/N]: " -ForegroundColor Yellow -NoNewline
    $skipInput = Read-Host
    if ($skipInput.Trim() -match '^[Yy]') {
        $SkipExisting = $true
    }

    # ── 7. Confirm ──
    Show-MenuHeader -Title "Confirm Bulk Move"
    Write-Host "  Source Collection: $SourceCollection" -ForegroundColor White
    Write-Host "  Source Project:    $SourceProject" -ForegroundColor White
    Write-Host "  Target Collection: $TargetCollection" -ForegroundColor White
    Write-Host "  Target Project:    $TargetProject" -ForegroundColor White
    Write-Host "  Repos to move:     $($selectedRepoNames.Count)" -ForegroundColor White
    Write-Host "  Skip existing:     $(if ($SkipExisting) { 'yes' } else { 'no (push into existing)' })" -ForegroundColor White
    Write-Host ""
    foreach ($name in $selectedRepoNames) {
        Write-Host "    - $name" -ForegroundColor DarkGray
    }
    Write-Host ""
    Write-Host "  Run as dry-run first (preview only, no changes)? [Y/n]: " -ForegroundColor Yellow -NoNewline
    $dryInput = Read-Host
    if (-not ($dryInput.Trim() -match '^[Nn]')) {
        $DryRun = $true
    }

    Write-Host "  Proceed? [Y/n]: " -ForegroundColor Yellow -NoNewline
    $confirm = Read-Host
    if ($confirm.Trim() -match '^[Nn]') {
        Write-Host "Cancelled." -ForegroundColor Yellow
        return
    }
}

# ─── Validate required params ──────────────────────────────────────────────────

if (-not $SourceCollection) { throw "SourceCollection is required. Use -Interactive or provide -SourceCollection." }
if (-not $SourceProject)    { throw "SourceProject is required. Use -Interactive or provide -SourceProject." }
if (-not $TargetCollection) { throw "TargetCollection is required. Use -Interactive or provide -TargetCollection." }
if (-not $TargetProject)    { throw "TargetProject is required. Use -Interactive or provide -TargetProject." }

if ($SourceCollection -eq $TargetCollection) {
    throw "Source and target collections must be different. Got '$SourceCollection' for both."
}

$sourceConfig = $config.collections[$SourceCollection]
$targetConfig = $config.collections[$TargetCollection]
if (-not $sourceConfig) { throw "Source collection '$SourceCollection' not in config." }
if (-not $targetConfig) { throw "Target collection '$TargetCollection' not in config." }

$sourcePat = $sourceConfig.pat
$targetPat = $targetConfig.pat

if (-not $sourcePat) { throw "Source collection '$SourceCollection' is missing a 'pat' in config." }
if (-not $targetPat) { throw "Target collection '$TargetCollection' is missing a 'pat' in config." }

# ─── Validate connections & project existence ──────────────────────────────────

Write-MigrationLog -Message "Starting bulk Git repo move" -LogFile $logFile -Level INFO
Write-MigrationLog -Message "  Source: $SourceCollection / $SourceProject" -LogFile $logFile
Write-MigrationLog -Message "  Target: $TargetCollection / $TargetProject" -LogFile $logFile

$sourceTest = Test-AdoConnection -ServerUrl $config.adoServerUrl -Collection $SourceCollection -Pat $sourcePat
if (-not $sourceTest.Connected) {
    throw "Cannot connect to source collection '$SourceCollection': $($sourceTest.Error)"
}
if ($SourceProject -notin $sourceTest.Projects) {
    throw "Project '$SourceProject' not found in collection '$SourceCollection'. Available: $($sourceTest.Projects -join ', ')"
}

$targetTest = Test-AdoConnection -ServerUrl $config.adoServerUrl -Collection $TargetCollection -Pat $targetPat
if (-not $targetTest.Connected) {
    throw "Cannot connect to target collection '$TargetCollection': $($targetTest.Error)"
}
if ($TargetProject -notin $targetTest.Projects) {
    throw "Project '$TargetProject' not found in collection '$TargetCollection'. Available: $($targetTest.Projects -join ', ')"
}

Write-MigrationLog -Message "Source and target connections verified" -LogFile $logFile -Level SUCCESS

# Resolve target project ID once
$targetProjects = Get-AdoProjects -ServerUrl $config.adoServerUrl -Collection $TargetCollection -Pat $targetPat
$targetProjectObj = $targetProjects | Where-Object { $_.name -eq $TargetProject }
if (-not $targetProjectObj) {
    throw "Project '$TargetProject' not found in '$TargetCollection' (project list lookup)."
}

# ─── Enumerate source repos & apply filters ────────────────────────────────────

if (-not $selectedRepoNames) {
    Write-MigrationLog -Message "Enumerating Git repos in source project..." -LogFile $logFile -Level INFO
    $sourceRepos = Get-AdoGitRepositories -ServerUrl $config.adoServerUrl `
        -Collection $SourceCollection -Pat $sourcePat -ProjectName $SourceProject
    if (-not $sourceRepos -or $sourceRepos.Count -eq 0) {
        Write-MigrationLog -Message "No Git repos found in $SourceCollection / $SourceProject." -LogFile $logFile -Level WARN
        return
    }
    $selectedRepoNames = @($sourceRepos | Sort-Object name | ForEach-Object { $_.name })
}
else {
    # Interactive mode already populated the list; still fetch the full repo objects
    $sourceRepos = Get-AdoGitRepositories -ServerUrl $config.adoServerUrl `
        -Collection $SourceCollection -Pat $sourcePat -ProjectName $SourceProject
}

# Apply Include/Exclude filters
$filteredNames = @()
foreach ($name in $selectedRepoNames) {
    if ($IncludeRepoNames -and -not (Test-RepoNameMatchesAny -Name $name -Patterns $IncludeRepoNames)) { continue }
    if ($ExcludeRepoNames -and (Test-RepoNameMatchesAny -Name $name -Patterns $ExcludeRepoNames)) { continue }
    $filteredNames += $name
}

if ($filteredNames.Count -eq 0) {
    Write-MigrationLog -Message "No repos remain after applying include/exclude filters." -LogFile $logFile -Level WARN
    return
}

$reposToMove = $sourceRepos | Where-Object { $_.name -in $filteredNames } | Sort-Object name
Write-MigrationLog -Message "Plan: move $($reposToMove.Count) repo(s) from $SourceCollection/$SourceProject to $TargetCollection/$TargetProject" -LogFile $logFile -Level INFO

# Pre-fetch the target repo list once so we can detect collisions without N API calls
$existingTargetRepos = Get-AdoGitRepositories -ServerUrl $config.adoServerUrl `
    -Collection $TargetCollection -Pat $targetPat -ProjectName $TargetProject
$existingTargetNames = @($existingTargetRepos | ForEach-Object { $_.name.ToLower() })

# ─── Dry run? Print plan and exit ──────────────────────────────────────────────

if ($DryRun) {
    Show-MenuHeader -Title "DRY RUN — no changes will be made"
    Write-Host "  $($reposToMove.Count) repo(s) would be moved:" -ForegroundColor White
    Write-Host ""
    foreach ($repo in $reposToMove) {
        $collides = $repo.name.ToLower() -in $existingTargetNames
        $note = if ($collides) {
            if ($SkipExisting) { 'EXISTS in target → would SKIP' } else { 'EXISTS in target → would PUSH into existing' }
        } else { 'new in target' }
        $size = if ($repo.size) { Format-Bytes -Bytes $repo.size } else { '?' }
        Write-Host ("    {0,-50} [{1,10}]  {2}" -f $repo.name, $size, $note) -ForegroundColor DarkGray
    }
    Write-Host ""
    Write-MigrationLog -Message "Dry run complete — no changes made." -LogFile $logFile -Level SUCCESS
    return
}

# ─── Prepare working directory ─────────────────────────────────────────────────

if (-not $WorkingDirectory) {
    $WorkingDirectory = Join-Path $config.outputDirectory "git-bulk-move/$SourceCollection-$SourceProject"
}
if (-not (Test-Path $WorkingDirectory)) {
    New-Item -ItemType Directory -Path $WorkingDirectory -Force | Out-Null
}
Write-MigrationLog -Message "Working directory: $WorkingDirectory" -LogFile $logFile -Level INFO

# ─── Move each repo ────────────────────────────────────────────────────────────

$results = [System.Collections.Generic.List[object]]::new()
$startTime = Get-Date
$totalCount = $reposToMove.Count
$index = 0

foreach ($repo in $reposToMove) {
    $index++
    $repoStart = Get-Date
    $pct = [int](($index / $totalCount) * 100)
    $elapsed = (Get-Date) - $startTime
    $eta = if ($index -gt 1) {
        $avgPerRepo = $elapsed.TotalSeconds / ($index - 1)
        $remaining = [TimeSpan]::FromSeconds($avgPerRepo * ($totalCount - $index + 1))
        " | ETA $([int]$remaining.TotalMinutes)m"
    } else { '' }

    Write-Host ""
    Write-Host "  [$index/$totalCount] ($pct%) $($repo.name)$eta" -ForegroundColor Cyan
    Write-MigrationLog -Message "[$index/$totalCount] Processing repo: $($repo.name)" -LogFile $logFile -Level INFO

    $result = [ordered]@{
        Name           = $repo.name
        SourceUrl      = $repo.remoteUrl
        TargetUrl      = $null
        Status         = $null
        Reason         = $null
        DurationSec    = 0
    }

    try {
        # ── Collision check ──
        $collides = $repo.name.ToLower() -in $existingTargetNames
        if ($collides -and $SkipExisting) {
            Write-MigrationLog -Message "  Skipping — already exists in target." -LogFile $logFile -Level WARN
            $result.Status = 'Skipped'
            $result.Reason = 'Already exists in target'
            continue
        }

        # ── Step 1: Mirror clone source ──
        $mirrorPath = Join-Path $WorkingDirectory "$($repo.name).git"
        if (Test-Path $mirrorPath) {
            Write-MigrationLog -Message "  Removing stale local clone: $mirrorPath" -LogFile $logFile -Level INFO
            Remove-Item -Path $mirrorPath -Recurse -Force
        }

        $sourceAuthUrl = $repo.remoteUrl -replace '://', "://:$sourcePat@"
        Write-MigrationLog -Message "  Step 1/3: Mirror cloning source" -LogFile $logFile -Level INFO
        Invoke-Git -Arguments "clone --mirror `"$sourceAuthUrl`" `"$mirrorPath`"" -LogFile $logFile

        # ── Step 2: Create (or fetch) target repo ──
        Write-MigrationLog -Message "  Step 2/3: Ensuring target repo exists" -LogFile $logFile -Level INFO
        $targetRepo = New-AdoGitRepository -ServerUrl $config.adoServerUrl `
            -Collection $TargetCollection -Pat $targetPat `
            -ProjectName $TargetProject -ProjectId $targetProjectObj.id `
            -RepoName $repo.name
        $result.TargetUrl = $targetRepo.remoteUrl
        Write-MigrationLog -Message "  Target: $($targetRepo.remoteUrl)" -LogFile $logFile -Level INFO

        # ── Step 3: Push mirror to target ──
        Write-MigrationLog -Message "  Step 3/3: Pushing mirror to target" -LogFile $logFile -Level INFO
        $targetAuthUrl = $targetRepo.remoteUrl -replace '://', "://:$targetPat@"
        Invoke-Git -Arguments "remote set-url --push origin `"$targetAuthUrl`"" -WorkingDirectory $mirrorPath -LogFile $logFile
        Invoke-Git -Arguments "push --mirror" -WorkingDirectory $mirrorPath -LogFile $logFile
        # Strip the embedded PAT from local config
        Invoke-Git -Arguments "remote set-url --push origin `"$($targetRepo.remoteUrl)`"" -WorkingDirectory $mirrorPath -LogFile $logFile

        # Refresh the existing-target list so a re-run of the same script sees this repo as existing
        $existingTargetNames += $repo.name.ToLower()

        if (-not $KeepWorkingCopies) {
            Remove-Item -Path $mirrorPath -Recurse -Force
        }

        $result.Status = if ($collides) { 'PushedIntoExisting' } else { 'Moved' }
        Write-MigrationLog -Message "  ✓ $($repo.name) → $($targetRepo.remoteUrl)" -LogFile $logFile -Level SUCCESS
    }
    catch {
        $result.Status = 'Failed'
        $result.Reason = $_.Exception.Message
        Write-MigrationLog -Message "  ✗ $($repo.name) failed: $($_.Exception.Message)" -LogFile $logFile -Level ERROR
    }
    finally {
        $result.DurationSec = [int]((Get-Date) - $repoStart).TotalSeconds
        $results.Add([pscustomobject]$result)
    }
}

# ─── Summary & manifest ────────────────────────────────────────────────────────

$totalElapsed = (Get-Date) - $startTime
$moved      = @($results | Where-Object { $_.Status -in @('Moved', 'PushedIntoExisting') })
$skipped    = @($results | Where-Object { $_.Status -eq 'Skipped' })
$failed     = @($results | Where-Object { $_.Status -eq 'Failed' })

$failedColor = if ($failed.Count -gt 0) { 'Red' } else { 'DarkGray' }

Show-MenuHeader -Title "Bulk Git Repo Move — Summary"
Write-Host "  Source:     $SourceCollection / $SourceProject" -ForegroundColor White
Write-Host "  Target:     $TargetCollection / $TargetProject" -ForegroundColor White
Write-Host "  Total:      $($results.Count)" -ForegroundColor White
Write-Host "  Moved:      $($moved.Count)" -ForegroundColor Green
Write-Host "  Skipped:    $($skipped.Count)" -ForegroundColor Yellow
Write-Host "  Failed:     $($failed.Count)" -ForegroundColor $failedColor
Write-Host "  Elapsed:    $([int]$totalElapsed.TotalMinutes)m $([int]($totalElapsed.TotalSeconds % 60))s" -ForegroundColor White

if ($failed.Count -gt 0) {
    Write-Host ""
    Write-Host "  Failures:" -ForegroundColor Red
    foreach ($f in $failed) {
        Write-Host "    - $($f.Name): $($f.Reason)" -ForegroundColor Red
    }
}

$manifestPath = Join-Path $config.logDirectory "git-bulk-move-$(Get-Date -Format 'yyyyMMdd-HHmmss').csv"
$results | Export-Csv -Path $manifestPath -NoTypeInformation -Encoding utf8
Write-Host ""
Write-Host "  Manifest:   $manifestPath" -ForegroundColor DarkGray
Write-Host "  Log:        $logFile" -ForegroundColor DarkGray
Write-Host ""

Write-MigrationLog -Message "Bulk move complete. Moved=$($moved.Count), Skipped=$($skipped.Count), Failed=$($failed.Count). Manifest: $manifestPath" -LogFile $logFile -Level SUCCESS

if ($failed.Count -gt 0) {
    exit 1
}
