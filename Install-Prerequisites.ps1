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

    If a tool is missing, prompts the user to install it automatically
    (using Chocolatey, winget, or pip where available).

.EXAMPLE
    ./Install-Prerequisites.ps1
#>

[CmdletBinding()]
param()

$ErrorActionPreference = 'Continue'

Write-Host ""
Write-Host "═══════════════════════════════════════════════════════" -ForegroundColor Cyan
Write-Host "  ADO TFVC-to-Git Migrator — Prerequisite Check" -ForegroundColor Cyan
Write-Host "═══════════════════════════════════════════════════════" -ForegroundColor Cyan
Write-Host ""

$allGood = $true
$missing = [System.Collections.ArrayList]::new()

# ─── Detect package managers ──────────────────────────────────────────────────

$hasChoco = $null -ne (Get-Command 'choco' -ErrorAction SilentlyContinue)
$hasWinget = $null -ne (Get-Command 'winget' -ErrorAction SilentlyContinue)
$hasPip = $null -ne (Get-Command 'pip' -ErrorAction SilentlyContinue)
if (-not $hasPip) { $hasPip = $null -ne (Get-Command 'pip3' -ErrorAction SilentlyContinue) }

# ─── PowerShell ────────────────────────────────────────────────────────────────

Write-Host "  Checking PowerShell..." -NoNewline
$psVersion = $PSVersionTable.PSVersion
if ($psVersion.Major -ge 7) {
    Write-Host " OK ($psVersion)" -ForegroundColor Green
}
else {
    Write-Host " FAIL (need 7+, have $psVersion)" -ForegroundColor Red
    Write-Host "    Since you're running this script, you likely already have PS7." -ForegroundColor Yellow
    Write-Host "    Make sure you're running 'pwsh' not 'powershell'." -ForegroundColor Yellow
    $allGood = $false
}

# ─── Git ──────────────────────────────────────────────────────────────────────

Write-Host "  Checking Git..." -NoNewline
$gitCmd = Get-Command 'git' -ErrorAction SilentlyContinue
if ($gitCmd) {
    $gitVersion = (git --version) -replace 'git version ', ''
    $major, $minor = $gitVersion.Split('.')[0..1] | ForEach-Object { [int]$_ }
    if ($major -gt 2 -or ($major -eq 2 -and $minor -ge 30)) {
        Write-Host " OK ($gitVersion)" -ForegroundColor Green
    }
    else {
        Write-Host " WARN (version $gitVersion, recommend 2.30+)" -ForegroundColor Yellow
        [void]$missing.Add('git')
        $allGood = $false
    }
}
else {
    Write-Host " NOT FOUND" -ForegroundColor Red
    [void]$missing.Add('git')
    $allGood = $false
}

# ─── git-tfs ──────────────────────────────────────────────────────────────────

Write-Host "  Checking git-tfs..." -NoNewline
$gitTfsFound = $false
$gitTfs = Get-Command 'git-tfs' -ErrorAction SilentlyContinue
if ($gitTfs) {
    $gitTfsFound = $true
}
else {
    try {
        $result = & git tfs --version 2>$null
        if ($result) { $gitTfsFound = $true }
    }
    catch { }
}

if ($gitTfsFound) {
    try {
        $version = & git tfs --version 2>$null
        Write-Host " OK ($version)" -ForegroundColor Green
    }
    catch {
        Write-Host " OK (found)" -ForegroundColor Green
    }
}
else {
    Write-Host " NOT FOUND" -ForegroundColor Red
    [void]$missing.Add('git-tfs')
    $allGood = $false
}

# ─── git-filter-repo ─────────────────────────────────────────────────────────

Write-Host "  Checking git-filter-repo..." -NoNewline
$filterRepo = Get-Command 'git-filter-repo' -ErrorAction SilentlyContinue
if ($filterRepo) {
    Write-Host " OK" -ForegroundColor Green
}
else {
    Write-Host " NOT FOUND (optional — needed for splitting repos)" -ForegroundColor Yellow
    [void]$missing.Add('git-filter-repo')
}

# ─── Git LFS ─────────────────────────────────────────────────────────────────

Write-Host "  Checking Git LFS..." -NoNewline
$gitLfs = Get-Command 'git-lfs' -ErrorAction SilentlyContinue
if ($gitLfs) {
    $lfsVersion = (git lfs version) -replace 'git-lfs/', '' -replace ' .*', ''
    Write-Host " OK ($lfsVersion)" -ForegroundColor Green
}
else {
    Write-Host " NOT FOUND (optional — needed for large binary files)" -ForegroundColor Yellow
    [void]$missing.Add('git-lfs')
}

# ─── Offer to install missing tools ──────────────────────────────────────────

function Find-OrInstall {
    <#
    .SYNOPSIS
        For a missing tool: check common paths, ask user to locate it, offer to add to PATH, or install via package manager.
    #>
    param(
        [string]$ToolName,
        [string]$ExeName,
        [string[]]$CommonPaths,
        [string]$Required,
        [scriptblock]$InstallAction
    )

    $label = if ($Required -eq 'required') { "(required)" } else { "(optional)" }
    Write-Host "  $ToolName $label" -ForegroundColor White

    # 1. Search common locations
    $foundPath = $null
    foreach ($p in $CommonPaths) {
        $candidate = Join-Path $p $ExeName
        if (Test-Path $candidate) {
            $foundPath = $p
            break
        }
    }

    if ($foundPath) {
        Write-Host "    Found at: $foundPath" -ForegroundColor Green
        Write-Host "    Add this to your PATH so it's always available? [Y/n]: " -ForegroundColor Yellow -NoNewline
        $answer = Read-Host
        if ($answer.Trim() -notmatch '^[Nn]') {
            Add-ToUserPath -Directory $foundPath
        }
        Write-Host ""
        return
    }

    # 2. Ask if they already have it downloaded somewhere
    Write-Host "    Do you already have $ToolName downloaded somewhere on this computer? [y/N]: " -ForegroundColor Yellow -NoNewline
    $hasIt = Read-Host
    if ($hasIt.Trim() -match '^[Yy]') {
        Write-Host "    Paste the folder path where $ExeName is located" -ForegroundColor Yellow
        Write-Host "    (e.g. E:\Tools\$ToolName): " -ForegroundColor DarkGray -NoNewline
        $userPath = Read-Host

        if ($userPath -and (Test-Path $userPath)) {
            $candidate = Join-Path $userPath.Trim() $ExeName
            if (Test-Path $candidate) {
                Write-Host "    Found $ExeName in $userPath" -ForegroundColor Green
                Write-Host "    Add this to your PATH? [Y/n]: " -ForegroundColor Yellow -NoNewline
                $answer = Read-Host
                if ($answer.Trim() -notmatch '^[Nn]') {
                    Add-ToUserPath -Directory $userPath.Trim()
                }
                Write-Host ""
                return
            }
            else {
                Write-Host "    Could not find '$ExeName' in that folder." -ForegroundColor Red
                # List what's there to help them
                $exes = Get-ChildItem $userPath.Trim() -Filter '*.exe' -ErrorAction SilentlyContinue | Select-Object -First 10
                if ($exes) {
                    Write-Host "    Files found there: $($exes.Name -join ', ')" -ForegroundColor DarkGray
                }
            }
        }
        else {
            Write-Host "    Folder not found: $userPath" -ForegroundColor Red
        }
    }

    # 3. Offer to install via package manager
    if ($InstallAction) {
        & $InstallAction
    }
    else {
        Write-Host "    No automatic installer available for $ToolName." -ForegroundColor DarkGray
    }
    Write-Host ""
}

function Add-ToUserPath {
    param([string]$Directory)
    $currentPath = [Environment]::GetEnvironmentVariable('PATH', 'User')
    if ($currentPath -split ';' | Where-Object { $_.TrimEnd('\') -eq $Directory.TrimEnd('\') }) {
        Write-Host "    Already in PATH." -ForegroundColor DarkGray
        return
    }
    [Environment]::SetEnvironmentVariable('PATH', "$currentPath;$Directory", 'User')
    # Also update current session so it works immediately
    $env:PATH = "$env:PATH;$Directory"
    Write-Host "    ✓ Added to PATH. It will persist across terminal restarts." -ForegroundColor Green
}

if ($missing.Count -gt 0) {
    Write-Host ""
    Write-Host "  ─── Missing Tools ─────────────────────────────────────" -ForegroundColor Yellow
    Write-Host ""

    foreach ($tool in $missing) {
        switch ($tool) {
            'git' {
                Find-OrInstall -ToolName 'Git' -ExeName 'git.exe' -Required 'required' `
                    -CommonPaths @(
                        'C:\Program Files\Git\cmd',
                        'C:\Program Files (x86)\Git\cmd',
                        'C:\Program Files\Git\bin'
                    ) `
                    -InstallAction {
                        if ($hasWinget) {
                            Write-Host "    Install with winget? [Y/n]: " -ForegroundColor Yellow -NoNewline
                            $a = Read-Host
                            if ($a.Trim() -notmatch '^[Nn]') {
                                Write-Host "    Installing Git..." -ForegroundColor Cyan
                                & winget install --id Git.Git -e --accept-package-agreements --accept-source-agreements
                                Write-Host "    Done. You may need to restart your terminal." -ForegroundColor Green
                            }
                        }
                        elseif ($hasChoco) {
                            Write-Host "    Install with Chocolatey? [Y/n]: " -ForegroundColor Yellow -NoNewline
                            $a = Read-Host
                            if ($a.Trim() -notmatch '^[Nn]') {
                                Write-Host "    Installing Git..." -ForegroundColor Cyan
                                & choco install git -y
                                Write-Host "    Done. You may need to restart your terminal." -ForegroundColor Green
                            }
                        }
                        else {
                            Write-Host "    Download manually: https://git-scm.com/downloads" -ForegroundColor Yellow
                        }
                    }
            }
            'git-tfs' {
                Find-OrInstall -ToolName 'git-tfs' -ExeName 'git-tfs.exe' -Required 'required' `
                    -CommonPaths @(
                        'C:\Program Files\git-tfs',
                        'C:\Program Files (x86)\git-tfs',
                        'C:\Tools\git-tfs',
                        'C:\ProgramData\chocolatey\lib\gittfs\tools'
                    ) `
                    -InstallAction {
                        if ($hasChoco) {
                            Write-Host "    Install with Chocolatey? [Y/n]: " -ForegroundColor Yellow -NoNewline
                            $a = Read-Host
                            if ($a.Trim() -notmatch '^[Nn]') {
                                Write-Host "    Installing git-tfs..." -ForegroundColor Cyan
                                & choco install gittfs -y
                                Write-Host "    Done. You may need to restart your terminal." -ForegroundColor Green
                            }
                        }
                        else {
                            Write-Host "    Download from: https://github.com/git-tfs/git-tfs/releases" -ForegroundColor Yellow
                            Write-Host "    Extract the zip, then re-run this check and point it to the folder." -ForegroundColor DarkGray
                        }
                        Write-Host "    NOTE: git-tfs requires .NET Framework (usually already on Windows)." -ForegroundColor DarkGray
                    }
            }
            'git-filter-repo' {
                Find-OrInstall -ToolName 'git-filter-repo' -ExeName 'git-filter-repo.exe' -Required 'optional' `
                    -CommonPaths @(
                        'C:\Python3*\Scripts',
                        'C:\Users\*\AppData\Local\Programs\Python\Python3*\Scripts'
                    ) `
                    -InstallAction {
                        $pipCmd = if (Get-Command 'pip3' -ErrorAction SilentlyContinue) { 'pip3' } elseif (Get-Command 'pip' -ErrorAction SilentlyContinue) { 'pip' } else { $null }
                        if ($pipCmd) {
                            Write-Host "    Install with pip? [Y/n]: " -ForegroundColor Yellow -NoNewline
                            $a = Read-Host
                            if ($a.Trim() -notmatch '^[Nn]') {
                                Write-Host "    Installing git-filter-repo..." -ForegroundColor Cyan
                                & $pipCmd install git-filter-repo
                                Write-Host "    Done." -ForegroundColor Green
                            }
                        }
                        elseif ($hasChoco) {
                            Write-Host "    Install with Chocolatey? [Y/n]: " -ForegroundColor Yellow -NoNewline
                            $a = Read-Host
                            if ($a.Trim() -notmatch '^[Nn]') {
                                Write-Host "    Installing git-filter-repo..." -ForegroundColor Cyan
                                & choco install git-filter-repo -y
                                Write-Host "    Done." -ForegroundColor Green
                            }
                        }
                        else {
                            Write-Host "    Install manually: https://github.com/newren/git-filter-repo" -ForegroundColor Yellow
                            Write-Host "    (The toolkit will fall back to git filter-branch if unavailable)" -ForegroundColor DarkGray
                        }
                    }
            }
            'git-lfs' {
                Find-OrInstall -ToolName 'Git LFS' -ExeName 'git-lfs.exe' -Required 'optional' `
                    -CommonPaths @(
                        'C:\Program Files\Git LFS',
                        'C:\Program Files\Git\mingw64\bin',
                        'C:\Program Files (x86)\Git LFS'
                    ) `
                    -InstallAction {
                        if ($hasWinget) {
                            Write-Host "    Install with winget? [Y/n]: " -ForegroundColor Yellow -NoNewline
                            $a = Read-Host
                            if ($a.Trim() -notmatch '^[Nn]') {
                                Write-Host "    Installing Git LFS..." -ForegroundColor Cyan
                                & winget install --id GitHub.GitLFS -e --accept-package-agreements --accept-source-agreements
                                & git lfs install 2>$null
                                Write-Host "    Done." -ForegroundColor Green
                            }
                        }
                        elseif ($hasChoco) {
                            Write-Host "    Install with Chocolatey? [Y/n]: " -ForegroundColor Yellow -NoNewline
                            $a = Read-Host
                            if ($a.Trim() -notmatch '^[Nn]') {
                                Write-Host "    Installing Git LFS..." -ForegroundColor Cyan
                                & choco install git-lfs -y
                                & git lfs install 2>$null
                                Write-Host "    Done." -ForegroundColor Green
                            }
                        }
                        else {
                            Write-Host "    Download manually: https://git-lfs.github.com/" -ForegroundColor Yellow
                        }
                    }
            }
        }
    }
}

# ─── Summary ──────────────────────────────────────────────────────────────────

Write-Host ""
Write-Host "═══════════════════════════════════════════════════════" -ForegroundColor Cyan
if ($allGood) {
    Write-Host "  All required prerequisites are installed!" -ForegroundColor Green
    Write-Host "  You're ready to run migrations." -ForegroundColor Cyan
}
else {
    Write-Host "  Some required tools were missing." -ForegroundColor Yellow
    Write-Host "  If you just installed them, restart your terminal and re-run this check." -ForegroundColor Yellow
}
Write-Host "═══════════════════════════════════════════════════════" -ForegroundColor Cyan
