#Requires -Version 7.0
<#
.SYNOPSIS
    Converts an entire TFVC repository (or branch path) to a Git repository using git-tfs.

.DESCRIPTION
    Uses git-tfs clone to convert a TFVC path to a full Git repo with history.
    Supports:
    - Full history or depth-limited conversion
    - Author mapping (DOMAIN\user → Git Name <email>)
    - Post-conversion cleanup (git-tfs metadata removal)
    - Git LFS setup for binary file types

.PARAMETER ConfigPath
    Path to migration-config.json.

.PARAMETER Collection
    ADO collection name.

.PARAMETER ProjectName
    Team project name.

.PARAMETER TfvcPath
    TFVC path to convert (e.g., $/MyProject or $/MyProject/Main).

.PARAMETER OutputRepoName
    Name for the output Git repository. Default: derived from TfvcPath.

.PARAMETER HistoryDepth
    Number of changesets to include. Null = full history.

.PARAMETER NoCleanup
    Skip post-conversion cleanup (keep git-tfs metadata).

.EXAMPLE
    ./Convert-TfvcToGit.ps1 -ConfigPath ./config/migration-config.json `
        -Collection GAMS -ProjectName "LegacyApp" -TfvcPath "$/LegacyApp"

.EXAMPLE
    ./Convert-TfvcToGit.ps1 -ConfigPath ./config/migration-config.json `
        -Collection GAMS -ProjectName "LegacyApp" -TfvcPath "$/LegacyApp/Main" `
        -OutputRepoName "legacy-app" -HistoryDepth 500
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory)]
    [string]$ConfigPath,

    [Parameter(Mandatory)]
    [string]$Collection,

    [Parameter(Mandatory)]
    [string]$ProjectName,

    [Parameter(Mandatory)]
    [string]$TfvcPath,

    [string]$OutputRepoName,

    [int]$HistoryDepth,

    [switch]$NoCleanup
)

$ErrorActionPreference = 'Stop'
Import-Module "$PSScriptRoot/modules/AdoTfvcMigrator.psm1" -Force

# ─── Load Config ───────────────────────────────────────────────────────────────

$config = Read-MigrationConfig -ConfigPath $ConfigPath
$logFile = Initialize-MigrationLog -LogDirectory $config.logDirectory -ScriptName 'ConvertTfvcToGit'

Write-MigrationLog -Message "Starting TFVC-to-Git conversion" -LogFile $logFile -Level INFO
Write-MigrationLog -Message "  Collection:  $Collection" -LogFile $logFile
Write-MigrationLog -Message "  Project:     $ProjectName" -LogFile $logFile
Write-MigrationLog -Message "  TFVC Path:   $TfvcPath" -LogFile $logFile

# ─── Validate ──────────────────────────────────────────────────────────────────

$collectionConfig = $config.collections[$Collection]
if (-not $collectionConfig) {
    throw "Collection '$Collection' not found in config."
}

$pat = $collectionConfig.pat

# Verify connection
$connTest = Test-AdoConnection -ServerUrl $config.adoServerUrl -Collection $Collection -Pat $pat
if (-not $connTest.Connected) {
    throw "Cannot connect to collection '$Collection': $($connTest.Error)"
}

# Verify git-tfs
$gitTfsPath = Find-GitTfs -GitTfsPath $config.gitTfsPath
Write-MigrationLog -Message "Using git-tfs: $gitTfsPath" -LogFile $logFile

# ─── Determine Output Path ────────────────────────────────────────────────────

if (-not $OutputRepoName) {
    # Derive from TFVC path: $/Project/Folder → project-folder
    $OutputRepoName = ($TfvcPath -replace '^\$/', '' -replace '/', '-').ToLower()
}

$outputPath = Join-Path $config.outputDirectory $OutputRepoName

if (Test-Path $outputPath) {
    Write-MigrationLog -Message "Output path already exists: $outputPath" -LogFile $logFile -Level WARN
    Write-MigrationLog -Message "Removing existing directory to start fresh" -LogFile $logFile -Level WARN
    Remove-Item -Recurse -Force $outputPath
}

# ─── Build git-tfs clone command ───────────────────────────────────────────────

$tfsUrl = "$($config.adoServerUrl)/$Collection"
$depth = if ($HistoryDepth) { $HistoryDepth } elseif ($config.migrationDefaults.historyDepthLimit) { $config.migrationDefaults.historyDepthLimit } else { $null }

$cloneArgs = "clone `"$tfsUrl`" `"$TfvcPath`" `"$outputPath`""

if ($depth) {
    $cloneArgs += " --changeset=$depth"
    Write-MigrationLog -Message "  History depth limited to $depth changesets" -LogFile $logFile
}

# Author mapping
$authorFile = $config.authorMappingFile
if ($authorFile -and (Test-Path $authorFile)) {
    $cloneArgs += " --authors=`"$authorFile`""
    Write-MigrationLog -Message "  Using author mapping: $authorFile" -LogFile $logFile
}

# Set PAT via environment variable for git-tfs authentication
$env:GIT_TFS_PAT = $pat

Write-MigrationLog -Message "Starting git-tfs clone (this may take a while)..." -LogFile $logFile -Level INFO

try {
    Invoke-GitTfs -Arguments $cloneArgs -LogFile $logFile
    Write-MigrationLog -Message "git-tfs clone completed successfully" -LogFile $logFile -Level SUCCESS
}
catch {
    Write-MigrationLog -Message "git-tfs clone failed: $($_.Exception.Message)" -LogFile $logFile -Level ERROR
    throw
}
finally {
    Remove-Item Env:\GIT_TFS_PAT -ErrorAction SilentlyContinue
}

# ─── Post-Conversion ──────────────────────────────────────────────────────────

# Rename default branch
$defaultBranch = $config.defaultBranch ?? 'main'
Write-MigrationLog -Message "Setting default branch to '$defaultBranch'" -LogFile $logFile

try {
    Invoke-Git -Arguments "branch -m master $defaultBranch" -WorkingDirectory $outputPath -LogFile $logFile
}
catch {
    Write-MigrationLog -Message "Branch rename skipped (may already be named correctly)" -LogFile $logFile -Level WARN
}

# Cleanup git-tfs metadata
if (-not $NoCleanup) {
    Remove-GitTfsMetadata -RepoPath $outputPath -LogFile $logFile
}

# Setup .gitignore
$gitignorePath = Join-Path $outputPath '.gitignore'
if (-not (Test-Path $gitignorePath)) {
    @"
# Build outputs
bin/
obj/
*.user
*.suo
*.cache
*.log

# Packages
packages/
node_modules/

# IDE
.vs/
*.swp
"@ | Out-File $gitignorePath -Encoding utf8

    Invoke-Git -Arguments "add .gitignore" -WorkingDirectory $outputPath -LogFile $logFile
    Invoke-Git -Arguments "commit -m `"Add .gitignore`"" -WorkingDirectory $outputPath -LogFile $logFile
}

# Git LFS for binaries
$lfsExtensions = $config.migrationDefaults.gitLfsExtensions
if ($lfsExtensions -and $lfsExtensions.Count -gt 0) {
    try {
        Remove-LargeFiles -RepoPath $outputPath -Extensions $lfsExtensions -LogFile $logFile
    }
    catch {
        Write-MigrationLog -Message "Git LFS setup skipped: $($_.Exception.Message)" -LogFile $logFile -Level WARN
    }
}

# ─── Summary ──────────────────────────────────────────────────────────────────

$commitCount = (Invoke-Git -Arguments "rev-list --count HEAD" -WorkingDirectory $outputPath -LogFile $logFile).Trim()
$repoSize = (Get-ChildItem -Recurse $outputPath | Measure-Object -Property Length -Sum).Sum / 1MB

Write-MigrationLog -Message "Conversion complete!" -LogFile $logFile -Level SUCCESS
Write-MigrationLog -Message "  Output:   $outputPath" -LogFile $logFile -Level SUCCESS
Write-MigrationLog -Message "  Commits:  $commitCount" -LogFile $logFile -Level SUCCESS
Write-MigrationLog -Message "  Size:     $([math]::Round($repoSize, 2)) MB" -LogFile $logFile -Level SUCCESS
Write-MigrationLog -Message "  Branch:   $defaultBranch" -LogFile $logFile -Level SUCCESS
Write-MigrationLog -Message "  Log:      $logFile" -LogFile $logFile -Level INFO

Write-Host ""
Write-Host "Next steps:" -ForegroundColor Cyan
Write-Host "  Push to GitHub:  ./Push-ToGitHub.ps1 -RepoPath `"$outputPath`" -GitHubOrg McDermott -GitHubRepo $OutputRepoName" -ForegroundColor Cyan
