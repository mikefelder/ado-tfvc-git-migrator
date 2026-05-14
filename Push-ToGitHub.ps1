#Requires -Version 7.0
<#
.SYNOPSIS
    Pushes a converted Git repository to GitHub Enterprise.

.DESCRIPTION
    Takes a local Git repo (output of Convert-TfvcToGit or Split-TfvcToGitRepos)
    and pushes it to a GitHub Enterprise organization.

    Supports an -Interactive mode that lists available converted repos in the
    output directory and lets you pick one to push.

    Steps:
    1. Creates the repo on GitHub Enterprise (via API)
    2. Sets the remote origin
    3. Pushes all branches and tags
    4. Optionally sets repo description, visibility, and default branch

.PARAMETER ConfigPath
    Path to migration-config.json (for GitHub settings). Optional if providing -GitHubUrl and -GitHubPat.

.PARAMETER Interactive
    Launch interactive mode — browse converted repos and select one to push.

.PARAMETER RepoPath
    Path to the local Git repository to push (skipped in interactive mode).

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
    # Interactive mode
    ./Push-ToGitHub.ps1 -ConfigPath ./config/migration-config.json -Interactive

.EXAMPLE
    ./Push-ToGitHub.ps1 -ConfigPath ./config/migration-config.json `
        -RepoPath ./output/legacy-app `
        -GitHubOrg "Contoso" -GitHubRepo "legacy-app"
#>

[CmdletBinding()]
param(
    [string]$ConfigPath,

    [switch]$Interactive,

    [string]$RepoPath,

    [string]$GitHubOrg,

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

# ─── Interactive Mode ──────────────────────────────────────────────────────────

if ($Interactive) {
    Show-MenuHeader -Title "Push Converted Repo to GitHub Enterprise"
    Write-Host "This wizard will help you select a converted Git repo and push it to GitHub." -ForegroundColor DarkGray

    # 1. Find converted repos in the output directory
    $outputDir = if ($config) { $config.outputDirectory } else { './output' }

    if (-not (Test-Path $outputDir)) {
        Write-Host ""
        Write-Host "  No output directory found at: $outputDir" -ForegroundColor Red
        Write-Host "  You need to convert a TFVC repo first (use Convert or Split from the main menu)." -ForegroundColor Yellow
        return
    }

    $gitRepos = @(Get-ChildItem $outputDir -Directory | Where-Object {
        Test-Path (Join-Path $_.FullName '.git')
    })

    if ($gitRepos.Count -eq 0) {
        Write-Host ""
        Write-Host "  No converted Git repos found in: $outputDir" -ForegroundColor Red
        Write-Host "  You need to convert a TFVC repo first (use Convert or Split from the main menu)." -ForegroundColor Yellow
        return
    }

    # Build display with repo info
    $displayItems = foreach ($repo in $gitRepos) {
        $commitCount = 'unknown'
        try {
            $commitCount = (& git -C $repo.FullName rev-list --count HEAD 2>$null).Trim()
        }
        catch { }
        $size = [math]::Round(($repo | Get-ChildItem -Recurse -File | Measure-Object -Property Length -Sum).Sum / 1MB, 1)
        "$($repo.Name)  ($commitCount commits, ${size} MB)"
    }

    Show-MenuHeader -Title "Select a repo to push to GitHub"
    $repoIdx = Show-NumberedMenu -Items $displayItems -Prompt 'Select a repo' -AllowBack
    if ($null -eq $repoIdx) { Write-Host "Cancelled." -ForegroundColor Yellow; return }

    $RepoPath = $gitRepos[$repoIdx].FullName
    $defaultRepoName = $gitRepos[$repoIdx].Name

    # 2. GitHub Org
    $defaultOrg = if ($config -and $config.github.defaultOrg) { $config.github.defaultOrg } else { 'Contoso' }
    Write-Host ""
    Write-Host "  GitHub organization [$defaultOrg]: " -ForegroundColor Yellow -NoNewline
    $orgInput = Read-Host
    $GitHubOrg = if ($orgInput.Trim()) { $orgInput.Trim() } else { $defaultOrg }

    # 3. GitHub Repo Name
    Write-Host "  GitHub repo name [$defaultRepoName]: " -ForegroundColor Yellow -NoNewline
    $repoNameInput = Read-Host
    $GitHubRepo = if ($repoNameInput.Trim()) { $repoNameInput.Trim() } else { $defaultRepoName }

    # 4. Description
    Write-Host "  Description [Migrated from ADO TFVC]: " -ForegroundColor Yellow -NoNewline
    $descInput = Read-Host
    if ($descInput.Trim()) { $Description = $descInput.Trim() }

    # 5. Visibility
    Write-Host "  Private repository? [Y/n]: " -ForegroundColor Yellow -NoNewline
    $visInput = Read-Host
    if ($visInput.Trim() -match '^[Nn]') { $Private = $false }

    # 6. Confirm
    Show-MenuHeader -Title "Confirm Push to GitHub"
    Write-Host "  Local Repo:    $RepoPath" -ForegroundColor White
    Write-Host "  GitHub Target: $GitHubUrl/$GitHubOrg/$GitHubRepo" -ForegroundColor White
    Write-Host "  Description:   $Description" -ForegroundColor White
    Write-Host "  Visibility:    $(if ($Private) { 'Private' } else { 'Public' })" -ForegroundColor White
    Write-Host ""
    Write-Host "  This will create the repo on GitHub and push all code and history." -ForegroundColor DarkGray
    Write-Host ""
    Write-Host "  Proceed? [Y/n]: " -ForegroundColor Yellow -NoNewline
    $confirm = Read-Host
    if ($confirm.Trim() -match '^[Nn]') {
        Write-Host "Cancelled." -ForegroundColor Yellow
        return
    }
}

# ─── Validate required params ─────────────────────────────────────────────────

if (-not $GitHubUrl) {
    Write-Host "  GitHub Enterprise URL is required." -ForegroundColor Red
    Write-Host "  Set it in your config file or pass -GitHubUrl." -ForegroundColor Yellow
    throw "GitHub Enterprise URL required."
}
if (-not $GitHubPat) {
    Write-Host "  GitHub PAT (Personal Access Token) is required." -ForegroundColor Red
    Write-Host "  Set it in your config file or pass -GitHubPat." -ForegroundColor Yellow
    throw "GitHub PAT required."
}
if (-not $RepoPath)   { throw "RepoPath is required. Use -Interactive or provide -RepoPath." }
if (-not $GitHubOrg)  { throw "GitHubOrg is required. Use -Interactive or provide -GitHubOrg." }
if (-not $GitHubRepo) { throw "GitHubRepo is required. Use -Interactive or provide -GitHubRepo." }

$logFile = Initialize-MigrationLog -LogDirectory $logDir -ScriptName 'PushToGitHub'

Write-MigrationLog -Message "Pushing to GitHub Enterprise" -LogFile $logFile -Level INFO
Write-MigrationLog -Message "  Source: $RepoPath" -LogFile $logFile
Write-MigrationLog -Message "  Target: $GitHubUrl/$GitHubOrg/$GitHubRepo" -LogFile $logFile

# ─── Validate Local Repo ──────────────────────────────────────────────────────

if (-not (Test-Path (Join-Path $RepoPath '.git'))) {
    Write-Host "  '$RepoPath' is not a Git repository." -ForegroundColor Red
    Write-Host "  Make sure you've converted the TFVC repo first." -ForegroundColor Yellow
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
