#Requires -Version 7.0
<#
.SYNOPSIS
    Checks and installs prerequisites for the TFVC-to-Git migration toolkit.

.DESCRIPTION
    Verifies that the following are available:
    - PowerShell 7+
    - Git 2.30+
    - git-tfs
    - git-filter-repo (optional but recommended)
    - Git LFS (optional)

    Provides download links and installation commands for missing tools.

.EXAMPLE
    ./Install-Prerequisites.ps1
#>

[CmdletBinding()]
param()

$ErrorActionPreference = 'Continue'

Write-Host "═══════════════════════════════════════════════════════" -ForegroundColor Cyan
Write-Host "  ADO TFVC-to-Git Migrator — Prerequisite Check" -ForegroundColor Cyan
Write-Host "═══════════════════════════════════════════════════════" -ForegroundColor Cyan
Write-Host ""

$allGood = $true

# ─── PowerShell ────────────────────────────────────────────────────────────────

Write-Host "Checking PowerShell..." -NoNewline
$psVersion = $PSVersionTable.PSVersion
if ($psVersion.Major -ge 7) {
    Write-Host " OK ($psVersion)" -ForegroundColor Green
}
else {
    Write-Host " FAIL (need 7+, have $psVersion)" -ForegroundColor Red
    Write-Host "  Install: https://learn.microsoft.com/powershell/scripting/install/installing-powershell" -ForegroundColor Yellow
    $allGood = $false
}

# ─── Git ──────────────────────────────────────────────────────────────────────

Write-Host "Checking Git..." -NoNewline
$gitCmd = Get-Command 'git' -ErrorAction SilentlyContinue
if ($gitCmd) {
    $gitVersion = (git --version) -replace 'git version ', ''
    $major, $minor = $gitVersion.Split('.')[0..1] | ForEach-Object { [int]$_ }
    if ($major -gt 2 -or ($major -eq 2 -and $minor -ge 30)) {
        Write-Host " OK ($gitVersion)" -ForegroundColor Green
    }
    else {
        Write-Host " WARN (version $gitVersion, recommend 2.30+)" -ForegroundColor Yellow
        $allGood = $false
    }
}
else {
    Write-Host " FAIL (not found)" -ForegroundColor Red
    Write-Host "  Install: https://git-scm.com/downloads" -ForegroundColor Yellow
    $allGood = $false
}

# ─── git-tfs ──────────────────────────────────────────────────────────────────

Write-Host "Checking git-tfs..." -NoNewline
$gitTfs = Get-Command 'git-tfs' -ErrorAction SilentlyContinue
if (-not $gitTfs) {
    # Also check as git subcommand
    try {
        $result = & git tfs --version 2>$null
        if ($result) { $gitTfs = $true }
    }
    catch { }
}

if ($gitTfs) {
    try {
        $version = & git tfs --version 2>$null
        Write-Host " OK ($version)" -ForegroundColor Green
    }
    catch {
        Write-Host " OK (found)" -ForegroundColor Green
    }
}
else {
    Write-Host " FAIL (not found)" -ForegroundColor Red
    Write-Host "  Install options:" -ForegroundColor Yellow
    Write-Host "    Chocolatey:  choco install gittfs" -ForegroundColor Yellow
    Write-Host "    Manual:      https://github.com/git-tfs/git-tfs/releases" -ForegroundColor Yellow
    Write-Host "    Scoop:       scoop install git-tfs" -ForegroundColor Yellow
    Write-Host ""
    Write-Host "  NOTE: git-tfs requires .NET Framework. On Windows, this is usually present." -ForegroundColor Yellow
    Write-Host "  For non-Windows, consider running this toolkit in a Windows VM or container." -ForegroundColor Yellow
    $allGood = $false
}

# ─── git-filter-repo ─────────────────────────────────────────────────────────

Write-Host "Checking git-filter-repo..." -NoNewline
$filterRepo = Get-Command 'git-filter-repo' -ErrorAction SilentlyContinue
if ($filterRepo) {
    Write-Host " OK" -ForegroundColor Green
}
else {
    Write-Host " MISSING (optional — needed for Split-TfvcToGitRepos)" -ForegroundColor Yellow
    Write-Host "  Install: pip install git-filter-repo" -ForegroundColor Yellow
    Write-Host "  Or:      https://github.com/newren/git-filter-repo" -ForegroundColor Yellow
    Write-Host "  (Falls back to git filter-branch if unavailable)" -ForegroundColor Yellow
}

# ─── Git LFS ─────────────────────────────────────────────────────────────────

Write-Host "Checking Git LFS..." -NoNewline
$gitLfs = Get-Command 'git-lfs' -ErrorAction SilentlyContinue
if ($gitLfs) {
    $lfsVersion = (git lfs version) -replace 'git-lfs/', '' -replace ' .*', ''
    Write-Host " OK ($lfsVersion)" -ForegroundColor Green
}
else {
    Write-Host " MISSING (optional — needed for binary file tracking)" -ForegroundColor Yellow
    Write-Host "  Install: https://git-lfs.github.com/" -ForegroundColor Yellow
    Write-Host "  Or:      choco install git-lfs" -ForegroundColor Yellow
}

# ─── Network Connectivity (optional) ─────────────────────────────────────────

Write-Host ""
Write-Host "Network connectivity test..." -NoNewline
Write-Host " SKIPPED (run Invoke-TfvcDiscovery.ps1 to test ADO connection)" -ForegroundColor Yellow

# ─── Summary ──────────────────────────────────────────────────────────────────

Write-Host ""
Write-Host "═══════════════════════════════════════════════════════" -ForegroundColor Cyan
if ($allGood) {
    Write-Host "  All required prerequisites are installed!" -ForegroundColor Green
    Write-Host "  Next: cp config/migration-config.example.json config/migration-config.json" -ForegroundColor Cyan
    Write-Host "  Then: Edit config/migration-config.json with your ADO server details" -ForegroundColor Cyan
}
else {
    Write-Host "  Some prerequisites are missing. Install them and re-run." -ForegroundColor Red
}
Write-Host "═══════════════════════════════════════════════════════" -ForegroundColor Cyan
