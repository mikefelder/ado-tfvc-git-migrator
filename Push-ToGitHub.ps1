#Requires -Version 7.0
<#
.SYNOPSIS
    Pushes a converted Git repository to GitHub Enterprise.

.DESCRIPTION
    Takes a local Git repo (output of Convert-TfvcToGit or Split-TfvcToGitRepos)
    and pushes it to a GitHub Enterprise organization.

    Steps:
    1. Creates the repo on GitHub Enterprise (via API)
    2. Sets the remote origin
    3. Pushes all branches and tags
    4. Optionally sets repo description, visibility, and default branch

.PARAMETER ConfigPath
    Path to migration-config.json (for GitHub settings). Optional if providing -GitHubUrl and -GitHubPat.

.PARAMETER RepoPath
    Path to the local Git repository to push.

.PARAMETER GitHubOrg
    GitHub organization name.

.PARAMETER GitHubRepo
    GitHub repository name.

.PARAMETER Description
    Repository description on GitHub.

.PARAMETER Private
    Create as a private repo. Default: true.

.PARAMETER GitHubUrl
    GitHub Enterprise URL. Overrides config.

.PARAMETER GitHubPat
    GitHub PAT. Overrides config.

.PARAMETER SkipRepoCreation
    Skip creating the GitHub repo (if it already exists).

.EXAMPLE
    ./Push-ToGitHub.ps1 -ConfigPath ./config/migration-config.json `
        -RepoPath ./output/legacy-app `
        -GitHubOrg "McDermott" -GitHubRepo "legacy-app"

.EXAMPLE
    ./Push-ToGitHub.ps1 -RepoPath ./output/service-a `
        -GitHubOrg "McDermott" -GitHubRepo "service-a" `
        -GitHubUrl "https://github.mcdermott.com" `
        -GitHubPat $env:GITHUB_PAT `
        -Description "Migrated from GAMS/$/MonoRepo/ServiceA"
#>

[CmdletBinding()]
param(
    [string]$ConfigPath,

    [Parameter(Mandatory)]
    [string]$RepoPath,

    [Parameter(Mandatory)]
    [string]$GitHubOrg,

    [Parameter(Mandatory)]
    [string]$GitHubRepo,

    [string]$Description = "Migrated from ADO TFVC",

    [bool]$Private = $true,

    [string]$GitHubUrl,

    [string]$GitHubPat,

    [switch]$SkipRepoCreation
)

$ErrorActionPreference = 'Stop'
Import-Module "$PSScriptRoot/modules/AdoTfvcMigrator.psm1" -Force

# ─── Resolve Config ───────────────────────────────────────────────────────────

$logDir = './logs'
if ($ConfigPath) {
    $config = Read-MigrationConfig -ConfigPath $ConfigPath
    $logDir = $config.logDirectory
    if (-not $GitHubUrl) { $GitHubUrl = $config.github.enterpriseUrl }
    if (-not $GitHubPat) { $GitHubPat = $config.github.pat }
    if (-not $GitHubOrg -and $config.github.defaultOrg) { $GitHubOrg = $config.github.defaultOrg }
}

if (-not $GitHubUrl) { throw "GitHub Enterprise URL required. Set in config or pass -GitHubUrl." }
if (-not $GitHubPat) { throw "GitHub PAT required. Set in config or pass -GitHubPat." }

$logFile = Initialize-MigrationLog -LogDirectory $logDir -ScriptName 'PushToGitHub'

Write-MigrationLog -Message "Pushing to GitHub Enterprise" -LogFile $logFile -Level INFO
Write-MigrationLog -Message "  Source: $RepoPath" -LogFile $logFile
Write-MigrationLog -Message "  Target: $GitHubUrl/$GitHubOrg/$GitHubRepo" -LogFile $logFile

# ─── Validate Local Repo ──────────────────────────────────────────────────────

if (-not (Test-Path (Join-Path $RepoPath '.git'))) {
    throw "Not a Git repository: $RepoPath"
}

# ─── Create GitHub Repo ───────────────────────────────────────────────────────

$apiUrl = if ($ConfigPath -and $config.github.apiUrl) {
    $config.github.apiUrl
} else {
    "$GitHubUrl/api/v3"
}

$headers = @{
    Authorization  = "token $GitHubPat"
    Accept         = 'application/vnd.github.v3+json'
    'Content-Type' = 'application/json'
}

if (-not $SkipRepoCreation) {
    Write-MigrationLog -Message "Creating GitHub repo: $GitHubOrg/$GitHubRepo" -LogFile $logFile -Level INFO

    $createBody = @{
        name        = $GitHubRepo
        description = $Description
        private     = $Private
        auto_init   = $false
    } | ConvertTo-Json

    try {
        $response = Invoke-RestMethod -Uri "$apiUrl/orgs/$GitHubOrg/repos" `
            -Headers $headers -Method POST -Body $createBody -ErrorAction Stop

        Write-MigrationLog -Message "GitHub repo created: $($response.html_url)" -LogFile $logFile -Level SUCCESS
    }
    catch {
        $statusCode = $_.Exception.Response.StatusCode
        if ($statusCode -eq 422 -or $statusCode -eq 409) {
            Write-MigrationLog -Message "Repo already exists on GitHub — will push to existing" -LogFile $logFile -Level WARN
        }
        else {
            Write-MigrationLog -Message "GitHub API error ($statusCode): $($_.Exception.Message)" -LogFile $logFile -Level ERROR
            throw
        }
    }
}

# ─── Set Remote & Push ─────────────────────────────────────────────────────────

$remoteUrl = "$GitHubUrl/$GitHubOrg/${GitHubRepo}.git"
$authenticatedUrl = $remoteUrl -replace '://', "://x-access-token:$GitHubPat@"

# Remove existing origin if present
$existingRemotes = Invoke-Git -Arguments "remote" -WorkingDirectory $RepoPath -LogFile $logFile
if ($existingRemotes -match 'origin') {
    Invoke-Git -Arguments "remote remove origin" -WorkingDirectory $RepoPath -LogFile $logFile
}

Invoke-Git -Arguments "remote add origin `"$authenticatedUrl`"" -WorkingDirectory $RepoPath -LogFile $logFile

Write-MigrationLog -Message "Pushing all branches and tags..." -LogFile $logFile -Level INFO

# Push all branches
Invoke-Git -Arguments "push origin --all --force" -WorkingDirectory $RepoPath -LogFile $logFile

# Push tags
Invoke-Git -Arguments "push origin --tags --force" -WorkingDirectory $RepoPath -LogFile $logFile

# Replace authenticated remote with clean URL
Invoke-Git -Arguments "remote set-url origin `"$remoteUrl`"" -WorkingDirectory $RepoPath -LogFile $logFile

# ─── Set Default Branch ───────────────────────────────────────────────────────

$defaultBranch = if ($ConfigPath -and $config.defaultBranch) { $config.defaultBranch } else { 'main' }

try {
    $updateBody = @{ default_branch = $defaultBranch } | ConvertTo-Json
    Invoke-RestMethod -Uri "$apiUrl/repos/$GitHubOrg/$GitHubRepo" `
        -Headers $headers -Method PATCH -Body $updateBody -ErrorAction Stop | Out-Null

    Write-MigrationLog -Message "Default branch set to '$defaultBranch'" -LogFile $logFile -Level SUCCESS
}
catch {
    Write-MigrationLog -Message "Could not set default branch: $($_.Exception.Message)" -LogFile $logFile -Level WARN
}

# ─── Summary ──────────────────────────────────────────────────────────────────

Write-Host ""
Write-MigrationLog -Message "Push complete!" -LogFile $logFile -Level SUCCESS
Write-MigrationLog -Message "  GitHub URL: $GitHubUrl/$GitHubOrg/$GitHubRepo" -LogFile $logFile -Level SUCCESS
Write-MigrationLog -Message "  Log:        $logFile" -LogFile $logFile -Level INFO
