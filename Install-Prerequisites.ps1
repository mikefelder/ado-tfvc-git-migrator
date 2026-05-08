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

    If required prerequisites are missing, offers to automatically download
    and install them, update PATH, and verify the installation.

.EXAMPLE
    ./Install-Prerequisites.ps1
#>

[CmdletBinding()]
param()

$ErrorActionPreference = 'Continue'

# ─── Helpers ───────────────────────────────────────────────────────────────────

$toolsDir = Join-Path $env:LOCALAPPDATA 'ado-tfvc-migrator-tools'

function Add-ToSessionPath {
    param([string]$Directory)
    if ($Directory -and (Test-Path $Directory) -and $env:PATH -notlike "*$Directory*") {
        $env:PATH = "$Directory;$env:PATH"
    }
}

function Add-ToUserPath {
    param([string]$Directory)
    if (-not (Test-Path $Directory)) { return }
    $currentPath = [Environment]::GetEnvironmentVariable('PATH', 'User')
    if ($currentPath -notlike "*$Directory*") {
        [Environment]::SetEnvironmentVariable('PATH', "$Directory;$currentPath", 'User')
        Write-Host "    Added $Directory to user PATH" -ForegroundColor DarkGray
    }
    Add-ToSessionPath $Directory
}

function Sync-ToolsDirToPath {
    # Scan toolsDir subdirectories for installed executables and add to session PATH
    if (-not (Test-Path $toolsDir)) { return }
    Get-ChildItem -Path $toolsDir -Directory -Recurse -ErrorAction SilentlyContinue | ForEach-Object {
        $hasExe = Get-ChildItem -Path $_.FullName -Filter '*.exe' -File -ErrorAction SilentlyContinue | Select-Object -First 1
        if ($hasExe) {
            Add-ToSessionPath $_.FullName
        }
    }
}

function Test-GitTfsAvailable {
    $cmd = Get-Command 'git-tfs' -ErrorAction SilentlyContinue
    if ($cmd) { return $true }
    try {
        $result = & git tfs --version 2>$null
        if ($result) { return $true }
    } catch { }
    # Fallback: check known install location directly
    $gitTfsDir = Join-Path $toolsDir 'git-tfs'
    if (Test-Path $gitTfsDir) {
        $exe = Get-ChildItem -Path $gitTfsDir -Recurse -Filter 'git-tfs.exe' -ErrorAction SilentlyContinue | Select-Object -First 1
        if ($exe) {
            Add-ToSessionPath $exe.DirectoryName
            return $true
        }
    }
    return $false
}

# ─── Check Functions ──────────────────────────────────────────────────────────

function Test-Prereq-Git {
    $gitCmd = Get-Command 'git' -ErrorAction SilentlyContinue
    if ($gitCmd) {
        $gitVersion = (git --version) -replace 'git version ', ''
        $major, $minor = $gitVersion.Split('.')[0..1] | ForEach-Object { [int]$_ }
        if ($major -gt 2 -or ($major -eq 2 -and $minor -ge 30)) {
            return @{ Status = 'OK'; Detail = $gitVersion }
        }
        return @{ Status = 'WARN'; Detail = "version $gitVersion, recommend 2.30+" }
    }
    return @{ Status = 'FAIL'; Detail = 'not found' }
}

function Test-Prereq-GitTfs {
    if (Test-GitTfsAvailable) {
        try {
            $version = & git tfs --version 2>$null
            return @{ Status = 'OK'; Detail = "$version" }
        } catch {
            return @{ Status = 'OK'; Detail = 'found' }
        }
    }
    return @{ Status = 'FAIL'; Detail = 'not found' }
}

function Test-Prereq-GitFilterRepo {
    $cmd = Get-Command 'git-filter-repo' -ErrorAction SilentlyContinue
    if ($cmd) {
        return @{ Status = 'OK'; Detail = 'found' }
    }
    return @{ Status = 'MISSING'; Detail = 'optional — requires Python; toolkit falls back to git filter-branch' }
}

function Test-Prereq-GitLfs {
    $cmd = Get-Command 'git-lfs' -ErrorAction SilentlyContinue
    if (-not $cmd) {
        # Check known install locations
        $searchPaths = @(
            "${env:ProgramFiles}\Git\cmd",
            "${env:ProgramFiles}\Git LFS",
            (Join-Path $toolsDir 'git-lfs')
        )
        foreach ($dir in $searchPaths) {
            if (Test-Path $dir) {
                $exe = Get-ChildItem -Path $dir -Recurse -Filter 'git-lfs.exe' -ErrorAction SilentlyContinue | Select-Object -First 1
                if ($exe) {
                    Add-ToSessionPath $exe.DirectoryName
                    $cmd = Get-Command 'git-lfs' -ErrorAction SilentlyContinue
                    break
                }
            }
        }
    }
    if ($cmd) {
        $gitCmd = Get-Command 'git' -ErrorAction SilentlyContinue
        if ($gitCmd) {
            try {
                $lfsVersion = (git lfs version) -replace 'git-lfs/', '' -replace ' .*', ''
                return @{ Status = 'OK'; Detail = $lfsVersion }
            } catch {
                return @{ Status = 'OK'; Detail = 'found (git not available for version check)' }
            }
        }
        return @{ Status = 'OK'; Detail = 'found (git not available for version check)' }
    }
    return @{ Status = 'MISSING'; Detail = 'optional — needed for binary file tracking' }
}

# ─── Install Functions ────────────────────────────────────────────────────────

function Install-Prereq-Git {
    $installed = $false

    # Try winget first
    $winget = Get-Command 'winget' -ErrorAction SilentlyContinue
    if ($winget) {
        Write-Host "  Trying winget..." -ForegroundColor Yellow
        $proc = Start-Process -FilePath 'winget' -ArgumentList 'install', '--id', 'Git.Git', '-e', '--source', 'winget', '--accept-package-agreements', '--accept-source-agreements' -Wait -PassThru -NoNewWindow
        if ($proc.ExitCode -eq 0) {
            $installed = $true
        } elseif ($proc.ExitCode -in @(-1978335189, -1978335191)) {
            Write-Host "    winget reports already installed" -ForegroundColor DarkGray
            $installed = $true
        } else {
            Write-Host "    winget failed (exit code $($proc.ExitCode)) — will try direct download" -ForegroundColor DarkGray
        }
    }

    # Fallback: download installer from GitHub
    if (-not $installed) {
        Write-Host "  Downloading Git installer from GitHub..." -ForegroundColor Yellow
        try {
            $releaseInfo = Invoke-RestMethod -Uri 'https://api.github.com/repos/git-for-windows/git/releases/latest' -Headers @{ 'User-Agent' = 'ado-tfvc-migrator' }
            $asset = $releaseInfo.assets | Where-Object { $_.name -match '64-bit\.exe$' -and $_.name -notmatch 'portable' } | Select-Object -First 1
            if (-not $asset) {
                $asset = $releaseInfo.assets | Where-Object { $_.name -match 'Git-.*-64-bit.*\.exe$' } | Select-Object -First 1
            }
            if (-not $asset) {
                Write-Host "  ERROR: Could not find Git installer in release assets" -ForegroundColor Red
                Write-Host "  Install manually: https://git-scm.com/downloads" -ForegroundColor Yellow
                return $false
            }

            $installerPath = Join-Path $env:TEMP $asset.name
            Write-Host "    Downloading $($asset.name)..." -ForegroundColor DarkGray
            Invoke-WebRequest -Uri $asset.browser_download_url -OutFile $installerPath -UseBasicParsing

            Write-Host "    Running silent install..." -ForegroundColor DarkGray
            $proc = Start-Process -FilePath $installerPath -ArgumentList '/VERYSILENT', '/NORESTART', '/NOCANCEL', '/SP-', '/CLOSEAPPLICATIONS', '/RESTARTAPPLICATIONS', '/COMPONENTS=icons,ext\reg\shellhere,assoc,assoc_sh' -Wait -PassThru
            Remove-Item $installerPath -Force -ErrorAction SilentlyContinue

            if ($proc.ExitCode -eq 0) {
                $installed = $true
            } else {
                Write-Host "  ERROR: Git installer exited with code $($proc.ExitCode)" -ForegroundColor Red
                return $false
            }
        }
        catch {
            Write-Host "  ERROR: Failed to download/install Git: $_" -ForegroundColor Red
            Write-Host "  Install manually: https://git-scm.com/downloads" -ForegroundColor Yellow
            return $false
        }
    }

    if ($installed) {
        # Add default Git install paths to session
        $gitPaths = @(
            "${env:ProgramFiles}\Git\cmd",
            "${env:ProgramFiles(x86)}\Git\cmd"
        )
        foreach ($p in $gitPaths) {
            Add-ToSessionPath $p
        }
    }
    return $installed
}

function Install-Prereq-GitTfs {
    Write-Host "  Installing git-tfs from GitHub releases..." -ForegroundColor Yellow
    $gitTfsDir = Join-Path $toolsDir 'git-tfs'
    if (-not (Test-Path $gitTfsDir)) {
        New-Item -ItemType Directory -Path $gitTfsDir -Force | Out-Null
    }

    try {
        # Get latest release URL from GitHub API
        $releaseInfo = Invoke-RestMethod -Uri 'https://api.github.com/repos/git-tfs/git-tfs/releases/latest' -Headers @{ 'User-Agent' = 'ado-tfvc-migrator' }
        $asset = $releaseInfo.assets | Where-Object { $_.name -like '*x64*' -or $_.name -like '*.zip' } | Select-Object -First 1
        if (-not $asset) {
            $asset = $releaseInfo.assets | Where-Object { $_.name -like '*.zip' } | Select-Object -First 1
        }
        if (-not $asset) {
            Write-Host "  ERROR: Could not find git-tfs release asset" -ForegroundColor Red
            return $false
        }

        $zipPath = Join-Path $env:TEMP "git-tfs-latest.zip"
        Write-Host "    Downloading $($asset.name)..." -ForegroundColor DarkGray
        Invoke-WebRequest -Uri $asset.browser_download_url -OutFile $zipPath -UseBasicParsing

        Write-Host "    Extracting to $gitTfsDir..." -ForegroundColor DarkGray
        Expand-Archive -Path $zipPath -DestinationPath $gitTfsDir -Force
        Remove-Item $zipPath -Force -ErrorAction SilentlyContinue

        # The zip may contain a nested folder — find git-tfs.exe
        $exePath = Get-ChildItem -Path $gitTfsDir -Recurse -Filter 'git-tfs.exe' | Select-Object -First 1
        if ($exePath) {
            $binDir = $exePath.DirectoryName
        } else {
            $binDir = $gitTfsDir
        }

        Add-ToUserPath $binDir
        return $true
    }
    catch {
        Write-Host "  ERROR: Failed to download/install git-tfs: $_" -ForegroundColor Red
        return $false
    }
}

function Install-Prereq-GitLfs {
    $installed = $false

    # Try winget first
    $winget = Get-Command 'winget' -ErrorAction SilentlyContinue
    if ($winget) {
        Write-Host "  Trying winget..." -ForegroundColor Yellow
        $proc = Start-Process -FilePath 'winget' -ArgumentList 'install', '--id', 'GitHub.GitLFS', '-e', '--source', 'winget', '--accept-package-agreements', '--accept-source-agreements' -Wait -PassThru -NoNewWindow
        # 0 = success, -1978335189 = no update (already installed), -1978335191 = already installed
        if ($proc.ExitCode -eq 0) {
            $installed = $true
        } elseif ($proc.ExitCode -in @(-1978335189, -1978335191)) {
            Write-Host "    winget reports already installed — will try direct download" -ForegroundColor DarkGray
        } else {
            Write-Host "    winget failed (exit code $($proc.ExitCode)) — will try direct download" -ForegroundColor DarkGray
        }
    }

    # Fallback: download from GitHub releases
    if (-not $installed) {
        Write-Host "  Downloading Git LFS from GitHub releases..." -ForegroundColor Yellow
        try {
            $releaseInfo = Invoke-RestMethod -Uri 'https://api.github.com/repos/git-lfs/git-lfs/releases/latest' -Headers @{ 'User-Agent' = 'ado-tfvc-migrator' }
            $asset = $releaseInfo.assets | Where-Object { $_.name -match 'windows-amd64.*\.zip$' } | Select-Object -First 1
            if (-not $asset) {
                $asset = $releaseInfo.assets | Where-Object { $_.name -match 'windows.*\.zip$' } | Select-Object -First 1
            }
            if (-not $asset) {
                Write-Host "  ERROR: Could not find Git LFS release asset for Windows" -ForegroundColor Red
                Write-Host "  Install manually: https://git-lfs.github.com/" -ForegroundColor Yellow
                return $false
            }

            $zipPath = Join-Path $env:TEMP "git-lfs-latest.zip"
            $lfsDir = Join-Path $toolsDir 'git-lfs'
            Write-Host "    Downloading $($asset.name)..." -ForegroundColor DarkGray
            Invoke-WebRequest -Uri $asset.browser_download_url -OutFile $zipPath -UseBasicParsing

            if (-not (Test-Path $lfsDir)) {
                New-Item -ItemType Directory -Path $lfsDir -Force | Out-Null
            }
            Write-Host "    Extracting to $lfsDir..." -ForegroundColor DarkGray
            Expand-Archive -Path $zipPath -DestinationPath $lfsDir -Force
            Remove-Item $zipPath -Force -ErrorAction SilentlyContinue

            # Find git-lfs.exe (may be in a nested folder)
            $exePath = Get-ChildItem -Path $lfsDir -Recurse -Filter 'git-lfs.exe' | Select-Object -First 1
            if ($exePath) {
                Add-ToUserPath $exePath.DirectoryName
                $installed = $true
            } else {
                Write-Host "  ERROR: git-lfs.exe not found in downloaded archive" -ForegroundColor Red
                return $false
            }
        }
        catch {
            Write-Host "  ERROR: Failed to download Git LFS: $_" -ForegroundColor Red
            Write-Host "  Install manually: https://git-lfs.github.com/" -ForegroundColor Yellow
            return $false
        }
    }

    if ($installed) {
        # Add common bundled paths to session
        foreach ($p in @("${env:ProgramFiles}\Git\cmd", "${env:ProgramFiles}\Git LFS")) {
            Add-ToSessionPath $p
        }
        # Initialize LFS
        try { & git lfs install 2>$null | Out-Null } catch { }
    }
    return $installed
}

# ═══════════════════════════════════════════════════════════════════════════════
# Main
# ═══════════════════════════════════════════════════════════════════════════════

Write-Host "═══════════════════════════════════════════════════════" -ForegroundColor Cyan
Write-Host "  ADO TFVC-to-Git Migrator — Prerequisite Check" -ForegroundColor Cyan
Write-Host "═══════════════════════════════════════════════════════" -ForegroundColor Cyan
Write-Host ""

# Ensure tools dir is in PATH for this session (picks up previous installs)
Add-ToSessionPath $toolsDir
Sync-ToolsDirToPath

$missingRequired = @()
$missingOptional = @()

# ─── PowerShell ────────────────────────────────────────────────────────────────

Write-Host "Checking PowerShell..." -NoNewline
$psVersion = $PSVersionTable.PSVersion
if ($psVersion.Major -ge 7) {
    Write-Host " OK ($psVersion)" -ForegroundColor Green
}
else {
    Write-Host " FAIL (need 7+, have $psVersion)" -ForegroundColor Red
    Write-Host "  Install: https://learn.microsoft.com/powershell/scripting/install/installing-powershell" -ForegroundColor Yellow
    # PowerShell can't auto-upgrade itself, so just warn
}

# ─── Git ──────────────────────────────────────────────────────────────────────

Write-Host "Checking Git..." -NoNewline
$gitCheck = Test-Prereq-Git
switch ($gitCheck.Status) {
    'OK'   { Write-Host " OK ($($gitCheck.Detail))" -ForegroundColor Green }
    'WARN' {
        Write-Host " WARN ($($gitCheck.Detail))" -ForegroundColor Yellow
        $missingRequired += 'Git'
    }
    'FAIL' {
        Write-Host " FAIL ($($gitCheck.Detail))" -ForegroundColor Red
        $missingRequired += 'Git'
    }
}

# ─── git-tfs ──────────────────────────────────────────────────────────────────

Write-Host "Checking git-tfs..." -NoNewline
$gitTfsCheck = Test-Prereq-GitTfs
switch ($gitTfsCheck.Status) {
    'OK'   { Write-Host " OK ($($gitTfsCheck.Detail))" -ForegroundColor Green }
    'FAIL' {
        Write-Host " FAIL ($($gitTfsCheck.Detail))" -ForegroundColor Red
        $missingRequired += 'git-tfs'
    }
}

# ─── git-filter-repo ─────────────────────────────────────────────────────────

Write-Host "Checking git-filter-repo..." -NoNewline
$filterRepoCheck = Test-Prereq-GitFilterRepo
switch ($filterRepoCheck.Status) {
    'OK'      { Write-Host " OK ($($filterRepoCheck.Detail))" -ForegroundColor Green }
    'MISSING' {
        Write-Host " MISSING ($($filterRepoCheck.Detail))" -ForegroundColor Yellow

    }
}

# ─── Git LFS ─────────────────────────────────────────────────────────────────

Write-Host "Checking Git LFS..." -NoNewline
$gitLfsCheck = Test-Prereq-GitLfs
switch ($gitLfsCheck.Status) {
    'OK'      { Write-Host " OK ($($gitLfsCheck.Detail))" -ForegroundColor Green }
    'MISSING' {
        Write-Host " MISSING ($($gitLfsCheck.Detail))" -ForegroundColor Yellow
        $missingOptional += 'Git-LFS'
    }
}

# ─── Network Connectivity (optional) ─────────────────────────────────────────

Write-Host ""
Write-Host "Network connectivity test..." -NoNewline
Write-Host " SKIPPED (run Invoke-TfvcDiscovery.ps1 to test ADO connection)" -ForegroundColor Yellow

# ─── Auto-Install Prompt ─────────────────────────────────────────────────────

$allMissing = $missingRequired + $missingOptional

if ($allMissing.Count -gt 0) {
    Write-Host ""
    Write-Host "═══════════════════════════════════════════════════════" -ForegroundColor Cyan
    if ($missingRequired.Count -gt 0) {
        Write-Host "  Missing required: $($missingRequired -join ', ')" -ForegroundColor Red
    }
    if ($missingOptional.Count -gt 0) {
        Write-Host "  Missing optional: $($missingOptional -join ', ')" -ForegroundColor Yellow
    }
    Write-Host "═══════════════════════════════════════════════════════" -ForegroundColor Cyan
    Write-Host ""

    $response = Read-Host "Would you like to automatically install the missing prerequisites? (Y/N)"

    if ($response -match '^[Yy]') {
        Write-Host ""
        Write-Host "─── Automatic Installation ─────────────────────────────" -ForegroundColor Cyan
        Write-Host ""

        $installResults = @{}

        foreach ($tool in $allMissing) {
            Write-Host "[$tool] Installing..." -ForegroundColor Cyan
            switch ($tool) {
                'Git'             { $installResults[$tool] = Install-Prereq-Git }
                'git-tfs'         { $installResults[$tool] = Install-Prereq-GitTfs }
                'Git-LFS'         { $installResults[$tool] = Install-Prereq-GitLfs }
            }
            if ($installResults[$tool]) {
                Write-Host "[$tool] Install completed." -ForegroundColor Green
            } else {
                Write-Host "[$tool] Install failed." -ForegroundColor Red
            }
            Write-Host ""
        }

        # ─── Refresh PATH and Re-verify ──────────────────────────────────
        Write-Host "─── Verifying Installations ────────────────────────────" -ForegroundColor Cyan
        Write-Host ""

        # Refresh PATH from registry to pick up changes made by installers
        $machinePath = [Environment]::GetEnvironmentVariable('PATH', 'Machine')
        $userPath = [Environment]::GetEnvironmentVariable('PATH', 'User')
        $env:PATH = "$userPath;$machinePath"
        # Re-add tools dir and scan subdirectories for installed executables
        Add-ToSessionPath $toolsDir
        Sync-ToolsDirToPath

        # Show which PATH entries were added
        Write-Host "PATH entries for installed tools:" -ForegroundColor DarkGray
        $userPathEntries = ($userPath -split ';') | Where-Object { $_ -like "*$toolsDir*" }
        if ($userPathEntries) {
            $userPathEntries | ForEach-Object { Write-Host "    [User PATH] $_" -ForegroundColor DarkGray }
        }
        foreach ($tool in $allMissing) {
            switch ($tool) {
                'Git' {
                    @("${env:ProgramFiles}\Git\cmd") | Where-Object { $env:PATH -like "*$_*" } | ForEach-Object {
                        Write-Host "    [Session PATH] $_" -ForegroundColor DarkGray
                    }
                }
                'Git-LFS' {
                    @("${env:ProgramFiles}\Git LFS") | Where-Object { $env:PATH -like "*$_*" } | ForEach-Object {
                        Write-Host "    [Session PATH] $_" -ForegroundColor DarkGray
                    }
                }
            }
        }
        Write-Host ""

        $verifyFailed = @()

        foreach ($tool in $allMissing) {
            Write-Host "Verifying $tool..." -NoNewline
            switch ($tool) {
                'Git' {
                    $check = Test-Prereq-Git
                    if ($check.Status -eq 'OK') {
                        Write-Host " OK ($($check.Detail))" -ForegroundColor Green
                    } else {
                        Write-Host " FAILED" -ForegroundColor Red
                        $verifyFailed += $tool
                    }
                }
                'git-tfs' {
                    $check = Test-Prereq-GitTfs
                    if ($check.Status -eq 'OK') {
                        Write-Host " OK ($($check.Detail))" -ForegroundColor Green
                    } else {
                        Write-Host " FAILED" -ForegroundColor Red
                        $verifyFailed += $tool
                    }
                }
                'Git-LFS' {
                    $check = Test-Prereq-GitLfs
                    if ($check.Status -eq 'OK') {
                        Write-Host " OK ($($check.Detail))" -ForegroundColor Green
                    } else {
                        Write-Host " FAILED" -ForegroundColor Red
                        $verifyFailed += $tool
                    }
                }
            }
        }

        Write-Host ""
        Write-Host "═══════════════════════════════════════════════════════" -ForegroundColor Cyan
        if ($verifyFailed.Count -eq 0) {
            Write-Host "  All prerequisites installed and verified!" -ForegroundColor Green
            Write-Host "  NOTE: You may need to restart your terminal for PATH" -ForegroundColor Yellow
            Write-Host "  changes to take effect in other sessions." -ForegroundColor Yellow
        } else {
            Write-Host "  Some tools could not be verified: $($verifyFailed -join ', ')" -ForegroundColor Red
            Write-Host "  Try restarting your terminal, or install them manually." -ForegroundColor Yellow
        }
        Write-Host "═══════════════════════════════════════════════════════" -ForegroundColor Cyan
    }
    else {
        Write-Host ""
        Write-Host "Manual installation references:" -ForegroundColor Yellow
        foreach ($tool in $allMissing) {
            switch ($tool) {
                'Git'             { Write-Host "  Git:             https://git-scm.com/downloads" -ForegroundColor Yellow }
                'git-tfs'         { Write-Host "  git-tfs:         https://github.com/git-tfs/git-tfs/releases" -ForegroundColor Yellow }
                'Git-LFS'         { Write-Host "  Git LFS:         https://git-lfs.github.com/" -ForegroundColor Yellow }
            }
        }
        Write-Host ""
        Write-Host "═══════════════════════════════════════════════════════" -ForegroundColor Cyan
        Write-Host "  Install the missing tools and re-run this script." -ForegroundColor Red
        Write-Host "═══════════════════════════════════════════════════════" -ForegroundColor Cyan
    }
}
else {
    # ─── All Good Summary ─────────────────────────────────────────────────────
    Write-Host ""
    Write-Host "═══════════════════════════════════════════════════════" -ForegroundColor Cyan
    Write-Host "  All required prerequisites are installed!" -ForegroundColor Green
    Write-Host "  Next: cp config/migration-config.example.json config/migration-config.json" -ForegroundColor Cyan
    Write-Host "  Then: Edit config/migration-config.json with your ADO server details" -ForegroundColor Cyan
    Write-Host "═══════════════════════════════════════════════════════" -ForegroundColor Cyan
}
