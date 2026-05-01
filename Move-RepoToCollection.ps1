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

    This is useful when consolidating or reorganizing before the GitHub migration.

.PARAMETER ConfigPath
    Path to migration-config.json.

.PARAMETER SourceCollection
    Source ADO collection name.

.PARAMETER SourceProject
    Source team project name.

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
    ./Move-RepoToCollection.ps1 -ConfigPath ./config/migration-config.json `
        -SourceCollection "GAMS" -SourceProject "LegacyApp" -TfvcPath "$/LegacyApp" `
        -TargetCollection "ModernApps" -TargetProject "Platform" -TargetRepoName "legacy-app"
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory)]
    [string]$ConfigPath,

    [Parameter(Mandatory)]
    [string]$SourceCollection,

    [Parameter(Mandatory)]
    [string]$SourceProject,

    [Parameter(Mandatory)]
    [string]$TfvcPath,

    [Parameter(Mandatory)]
    [string]$TargetCollection,

    [Parameter(Mandatory)]
    [string]$TargetProject,

    [Parameter(Mandatory)]
    [string]$TargetRepoName,

    [int]$HistoryDepth,

    [switch]$SkipTargetRepoCreation
)

$ErrorActionPreference = 'Stop'
Import-Module "$PSScriptRoot/modules/AdoTfvcMigrator.psm1" -Force

# ─── Load Config ───────────────────────────────────────────────────────────────

$config = Read-MigrationConfig -ConfigPath $ConfigPath
$logFile = Initialize-MigrationLog -LogDirectory $config.logDirectory -ScriptName 'MoveRepoToCollection'

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
