#Requires -Version 7.0
<#
.SYNOPSIS
    Moves or clones a TFVC repo from one ADO collection/project to a different collection/project as a Git repo.

.DESCRIPTION
    Handles the scenario where legacy repos need to be reorganized across ADO collections.
    Process:
    1. Converts the TFVC source to Git (via git-tfs)
    2. Creates a new Git repo in the target ADO collection/project
    3. Pushes the converted repo to the target

    Supports an -Interactive mode that walks you through selecting the source repo/folder
    and the target collection/project via numbered menus.

.PARAMETER ConfigPath
    Path to migration-config.json.

.PARAMETER Interactive
    Launch interactive mode — browse and select source and target via menus.

.PARAMETER SourceCollection
    Source ADO collection name (skipped in interactive mode).

.PARAMETER SourceProject
    Source team project name (skipped in interactive mode).

.PARAMETER TfvcPath
    TFVC path to migrate (e.g., $/SourceProject/AppFolder).

.PARAMETER TargetCollection
    Target ADO collection name.

.PARAMETER TargetProject
    Target team project name.

.PARAMETER TargetRepoName
    Name for the new Git repo in the target project.

.PARAMETER HistoryDepth
    Number of changesets to include. Null = full history.

.PARAMETER SkipTargetRepoCreation
    Skip creating the target repo (if it already exists).

.EXAMPLE
    # Interactive mode — guided menus
    ./Move-RepoToCollection.ps1 -ConfigPath ./config/migration-config.json -Interactive

.EXAMPLE
    # Direct mode — all parameters specified
    ./Move-RepoToCollection.ps1 -ConfigPath ./config/migration-config.json `
        -SourceCollection "GAMS" -SourceProject "LegacyApp" -TfvcPath "$/LegacyApp" `
        -TargetCollection "ModernApps" -TargetProject "Platform" -TargetRepoName "legacy-app"
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory)]
    [string]$ConfigPath,

    [switch]$Interactive,

    [string]$SourceCollection,
    [string]$SourceProject,
    [string]$TfvcPath,
    [string]$TargetCollection,
    [string]$TargetProject,
    [string]$TargetRepoName,

    [int]$HistoryDepth,

    [switch]$SkipTargetRepoCreation
)

$ErrorActionPreference = 'Stop'
Import-Module "$PSScriptRoot/modules/AdoTfvcMigrator.psm1" -Force

# ─── Load Config ───────────────────────────────────────────────────────────────

$config = Read-MigrationConfig -ConfigPath $ConfigPath
$logFile = Initialize-MigrationLog -LogDirectory $config.logDirectory -ScriptName 'MoveRepoToCollection'

# ─── Interactive Mode ──────────────────────────────────────────────────────────

if ($Interactive) {
    Show-MenuHeader -Title "Move TFVC Repo to Another Collection"
    Write-Host "This wizard will walk you through selecting a source TFVC repo/folder" -ForegroundColor DarkGray
    Write-Host "and a destination collection/project to move it to." -ForegroundColor DarkGray

    # ── 1. Pick source collection ──
    $SourceCollection = Select-AdoCollection -Config $config -Prompt 'Select SOURCE collection'
    if (-not $SourceCollection) { Write-Host "Cancelled." -ForegroundColor Yellow; return }

    # ── 2. Pick source project ──
    $SourceProject = Select-AdoProject -Config $config -Collection $SourceCollection -Prompt 'Select SOURCE project'
    if (-not $SourceProject) { Write-Host "Cancelled." -ForegroundColor Yellow; return }

    # ── 3. Browse and pick a TFVC folder ──
    $currentPath = "`$/$SourceProject"
    while ($true) {
        $selected = Select-TfvcFolders -Config $config -Collection $SourceCollection `
            -ProjectName $SourceProject -ParentPath $currentPath -Prompt 'Select a folder to move (or drill deeper)'
        if (-not $selected) { Write-Host "Cancelled." -ForegroundColor Yellow; return }

        $chosenPath = $selected[0]
        Write-Host ""
        Write-Host "  Selected: $chosenPath" -ForegroundColor Green
        Write-Host ""
        Write-Host "  [1] Use this folder as the source" -ForegroundColor White
        Write-Host "  [2] Drill into this folder to pick a subfolder" -ForegroundColor White
        Write-Host "  [0] Back" -ForegroundColor DarkGray
        Write-Host ""
        Write-Host "Choice: " -ForegroundColor Yellow -NoNewline
        $drillChoice = Read-Host

        switch ($drillChoice.Trim()) {
            '1' { $TfvcPath = $chosenPath; break }
            '2' { $currentPath = $chosenPath; continue }
            '0' { Write-Host "Cancelled." -ForegroundColor Yellow; return }
            default { $TfvcPath = $chosenPath; break }
        }
        if ($TfvcPath) { break }
    }

    # ── 4. Pick target collection ──
    Show-MenuHeader -Title "Select DESTINATION"
    Write-Host "  Source: $SourceCollection / $SourceProject / $TfvcPath" -ForegroundColor DarkGray
    Write-Host ""

    $TargetCollection = Select-AdoCollection -Config $config -Prompt 'Select TARGET collection'
    if (-not $TargetCollection) { Write-Host "Cancelled." -ForegroundColor Yellow; return }

    # ── 5. Pick target project ──
    $TargetProject = Select-AdoProject -Config $config -Collection $TargetCollection -Prompt 'Select TARGET project'
    if (-not $TargetProject) { Write-Host "Cancelled." -ForegroundColor Yellow; return }

    # ── 6. Repo name ──
    $defaultName = ($TfvcPath -replace '^\$/', '' -replace '/', '-').ToLower()
    Write-Host ""
    Write-Host "  Git repo name [$defaultName]: " -ForegroundColor Yellow -NoNewline
    $nameInput = Read-Host
    $TargetRepoName = if ($nameInput.Trim()) { $nameInput.Trim() } else { $defaultName }

    # ── 7. History depth ──
    Write-Host "  History depth (enter for full history, or a number): " -ForegroundColor Yellow -NoNewline
    $depthInput = Read-Host
    if ($depthInput.Trim() -match '^\d+$') {
        $HistoryDepth = [int]$depthInput.Trim()
    }

    # ── Confirm ──
    Show-MenuHeader -Title "Confirm Move"
    Write-Host "  Source Collection: $SourceCollection" -ForegroundColor White
    Write-Host "  Source Project:    $SourceProject" -ForegroundColor White
    Write-Host "  TFVC Path:         $TfvcPath" -ForegroundColor White
    Write-Host "  Target Collection: $TargetCollection" -ForegroundColor White
    Write-Host "  Target Project:    $TargetProject" -ForegroundColor White
    Write-Host "  Target Repo Name:  $TargetRepoName" -ForegroundColor White
    if ($HistoryDepth) {
        Write-Host "  History Depth:     $HistoryDepth changesets" -ForegroundColor White
    }
    else {
        Write-Host "  History Depth:     Full" -ForegroundColor White
    }
    Write-Host ""
    Write-Host "  Proceed? [Y/n]: " -ForegroundColor Yellow -NoNewline
    $confirm = Read-Host
    if ($confirm.Trim() -match '^[Nn]') {
        Write-Host "Cancelled." -ForegroundColor Yellow
        return
    }
}

# ─── Validate required params in non-interactive mode ──────────────────────────

if (-not $SourceCollection) { throw "SourceCollection is required. Use -Interactive or provide -SourceCollection." }
if (-not $SourceProject)    { throw "SourceProject is required. Use -Interactive or provide -SourceProject." }
if (-not $TfvcPath)         { throw "TfvcPath is required. Use -Interactive or provide -TfvcPath." }
if (-not $TargetCollection) { throw "TargetCollection is required. Use -Interactive or provide -TargetCollection." }
if (-not $TargetProject)    { throw "TargetProject is required. Use -Interactive or provide -TargetProject." }
if (-not $TargetRepoName)   { throw "TargetRepoName is required. Use -Interactive or provide -TargetRepoName." }

Write-MigrationLog -Message "Starting cross-collection move" -LogFile $logFile -Level INFO
Write-MigrationLog -Message "  Source: $SourceCollection/$SourceProject ($TfvcPath)" -LogFile $logFile
Write-MigrationLog -Message "  Target: $TargetCollection/$TargetProject/$TargetRepoName" -LogFile $logFile

# ─── Validate Source & Target ──────────────────────────────────────────────────

$sourceConfig = $config.collections[$SourceCollection]
$targetConfig = $config.collections[$TargetCollection]

if (-not $sourceConfig) { throw "Source collection '$SourceCollection' not in config." }
if (-not $targetConfig) { throw "Target collection '$TargetCollection' not in config." }

$sourcePat = $sourceConfig.pat
$targetPat = $targetConfig.pat

# Test source
$sourceTest = Test-AdoConnection -ServerUrl $config.adoServerUrl -Collection $SourceCollection -Pat $sourcePat
if (-not $sourceTest.Connected) {
    throw "Cannot connect to source collection '$SourceCollection': $($sourceTest.Error)"
}

# Test target
$targetTest = Test-AdoConnection -ServerUrl $config.adoServerUrl -Collection $TargetCollection -Pat $targetPat
if (-not $targetTest.Connected) {
    throw "Cannot connect to target collection '$TargetCollection': $($targetTest.Error)"
}

if ($TargetProject -notin $targetTest.Projects) {
    throw "Project '$TargetProject' not found in collection '$TargetCollection'. Available: $($targetTest.Projects -join ', ')"
}

Write-MigrationLog -Message "Source and target connections verified" -LogFile $logFile -Level SUCCESS

# ─── Step 1: Convert TFVC to Git ──────────────────────────────────────────────

Write-MigrationLog -Message "Step 1: Converting TFVC to Git" -LogFile $logFile -Level INFO

$convertParams = @{
    ConfigPath     = $ConfigPath
    Collection     = $SourceCollection
    ProjectName    = $SourceProject
    TfvcPath       = $TfvcPath
    OutputRepoName = $TargetRepoName
}
if ($HistoryDepth) {
    $convertParams.HistoryDepth = $HistoryDepth
}

& "$PSScriptRoot/Convert-TfvcToGit.ps1" @convertParams

$localRepoPath = Join-Path $config.outputDirectory $TargetRepoName

if (-not (Test-Path (Join-Path $localRepoPath '.git'))) {
    throw "Conversion failed — no Git repo found at $localRepoPath"
}

Write-MigrationLog -Message "Conversion complete" -LogFile $logFile -Level SUCCESS

# ─── Step 2: Create target Git repo in ADO ─────────────────────────────────────

if (-not $SkipTargetRepoCreation) {
    Write-MigrationLog -Message "Step 2: Creating Git repo '$TargetRepoName' in $TargetCollection/$TargetProject" -LogFile $logFile -Level INFO

    # Find the project ID
    $projects = Get-AdoProjects -ServerUrl $config.adoServerUrl -Collection $TargetCollection -Pat $targetPat
    $targetProjectObj = $projects | Where-Object { $_.name -eq $TargetProject }
    if (-not $targetProjectObj) {
        throw "Project '$TargetProject' not found in '$TargetCollection'."
    }

    $createBody = @{
        name    = $TargetRepoName
        project = @{
            id = $targetProjectObj.id
        }
    }

    try {
        $url = "$($config.adoServerUrl)/$TargetCollection/_apis/git/repositories"
        $newRepo = Invoke-AdoApi -Url $url -Pat $targetPat -Method POST -Body $createBody
        $remoteUrl = $newRepo.remoteUrl
        Write-MigrationLog -Message "Created repo: $remoteUrl" -LogFile $logFile -Level SUCCESS
    }
    catch {
        if ($_.Exception.Message -like '*already exists*' -or $_.Exception.Message -like '*409*') {
            Write-MigrationLog -Message "Repo already exists — will push to existing" -LogFile $logFile -Level WARN
            # Fetch existing repo URL
            $url = "$($config.adoServerUrl)/$TargetCollection/$TargetProject/_apis/git/repositories/$TargetRepoName"
            $existingRepo = Invoke-AdoApi -Url $url -Pat $targetPat
            $remoteUrl = $existingRepo.remoteUrl
        }
        else {
            throw
        }
    }
}
else {
    # Get existing repo URL
    $url = "$($config.adoServerUrl)/$TargetCollection/$TargetProject/_apis/git/repositories/$TargetRepoName"
    $existingRepo = Invoke-AdoApi -Url $url -Pat $targetPat
    $remoteUrl = $existingRepo.remoteUrl
}

# ─── Step 3: Push to target ────────────────────────────────────────────────────

Write-MigrationLog -Message "Step 3: Pushing to target repo" -LogFile $logFile -Level INFO

# Construct authenticated URL (PAT embedded for push)
$authenticatedUrl = $remoteUrl -replace '://', "://:$targetPat@"

# Add remote and push
Invoke-Git -Arguments "remote add target `"$authenticatedUrl`"" -WorkingDirectory $localRepoPath -LogFile $logFile

$defaultBranch = $config.defaultBranch ?? 'main'
Invoke-Git -Arguments "push target $defaultBranch --force" -WorkingDirectory $localRepoPath -LogFile $logFile
Invoke-Git -Arguments "push target --tags" -WorkingDirectory $localRepoPath -LogFile $logFile

# Remove remote with PAT from local config
Invoke-Git -Arguments "remote remove target" -WorkingDirectory $localRepoPath -LogFile $logFile

Write-MigrationLog -Message "Push complete!" -LogFile $logFile -Level SUCCESS

# ─── Summary ──────────────────────────────────────────────────────────────────

Write-Host ""
Write-MigrationLog -Message "Cross-collection move complete!" -LogFile $logFile -Level SUCCESS
Write-MigrationLog -Message "  Source:     $SourceCollection/$SourceProject ($TfvcPath)" -LogFile $logFile -Level SUCCESS
Write-MigrationLog -Message "  Target:     $TargetCollection/$TargetProject/$TargetRepoName" -LogFile $logFile -Level SUCCESS
Write-MigrationLog -Message "  Remote URL: $remoteUrl" -LogFile $logFile -Level SUCCESS
Write-MigrationLog -Message "  Local copy: $localRepoPath" -LogFile $logFile -Level INFO
Write-MigrationLog -Message "  Log:        $logFile" -LogFile $logFile -Level INFO
