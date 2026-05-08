#Requires -Version 7.0
<#
.SYNOPSIS
    Main launcher for the ADO TFVC-to-Git Migration Toolkit.

.DESCRIPTION
    Provides a friendly, menu-driven interface for all migration operations.
    No parameters needed — just run it and follow the prompts.

.PARAMETER ConfigPath
    Path to migration-config.json. Default: ./config/migration-config.json

.EXAMPLE
    ./Start-Menu.ps1
#>

[CmdletBinding()]
param(
    [string]$ConfigPath = (Join-Path $PSScriptRoot 'config/migration-config.json')
)

$ErrorActionPreference = 'Stop'
Import-Module "$PSScriptRoot/modules/AdoTfvcMigrator.psm1" -Force

# ─── Helpers ───────────────────────────────────────────────────────────────────

function Show-Banner {
    Clear-Host
    Write-Host ""
    Write-Host "  ╔═══════════════════════════════════════════════════════════╗" -ForegroundColor Cyan
    Write-Host "  ║                                                           ║" -ForegroundColor Cyan
    Write-Host "  ║     ADO TFVC → Git Migration Toolkit                     ║" -ForegroundColor Cyan
    Write-Host "  ║     McDermott — GitHub Enterprise Migration               ║" -ForegroundColor Cyan
    Write-Host "  ║                                                           ║" -ForegroundColor Cyan
    Write-Host "  ╚═══════════════════════════════════════════════════════════╝" -ForegroundColor Cyan
    Write-Host ""
}

function Show-ConfigStatus {
    if (Test-Path $ConfigPath) {
        try {
            $cfg = Read-MigrationConfig -ConfigPath $ConfigPath
            $collCount = $cfg.collections.Keys.Count
            $server = $cfg.adoServerUrl
            Write-Host "  Config:  " -NoNewline -ForegroundColor DarkGray
            Write-Host "Loaded ($collCount collection(s) — $server)" -ForegroundColor Green
        }
        catch {
            Write-Host "  Config:  " -NoNewline -ForegroundColor DarkGray
            Write-Host "Found but has errors — run Setup Wizard to fix" -ForegroundColor Yellow
        }
    }
    else {
        Write-Host "  Config:  " -NoNewline -ForegroundColor DarkGray
        Write-Host "Not found — run Setup Wizard first (option 1)" -ForegroundColor Red
    }
    Write-Host ""
}

function Pause-ForUser {
    Write-Host ""
    Write-Host "  Press Enter to return to the main menu..." -ForegroundColor DarkGray -NoNewline
    # Flush any buffered input so the user doesn't skip past this prompt
    while ([Console]::KeyAvailable) { [void][Console]::ReadKey($true) }
    Read-Host
}

function Test-ConfigReady {
    if (-not (Test-Path $ConfigPath)) {
        Write-Host ""
        Write-Host "  Configuration file not found." -ForegroundColor Red
        Write-Host "  Please run the Setup Wizard first (option 1 on the main menu)." -ForegroundColor Yellow
        Write-Host ""
        Pause-ForUser
        return $false
    }
    return $true
}

# ─── Main Menu Loop ───────────────────────────────────────────────────────────

while ($true) {
    Show-Banner
    Show-ConfigStatus

    Write-Host "  What would you like to do?" -ForegroundColor White
    Write-Host ""
    Write-Host "  ── Getting Started ────────────────────────────────────" -ForegroundColor DarkGray
    Write-Host "  [1]  Setup Wizard          Set up your server connection & settings" -ForegroundColor White
    Write-Host "  [2]  Check Prerequisites   Make sure all required tools are installed" -ForegroundColor White
    Write-Host ""
    Write-Host "  ── Explore ────────────────────────────────────────────" -ForegroundColor DarkGray
    Write-Host "  [3]  Discover Repos        Scan the server and list all repositories" -ForegroundColor White
    Write-Host ""
    Write-Host "  ── Migrate (one at a time) ────────────────────────────" -ForegroundColor DarkGray
    Write-Host "  [4]  Convert Repo          Convert a single repo from TFVC to Git" -ForegroundColor White
    Write-Host "  [5]  Split Repo            Break one large repo into smaller Git repos" -ForegroundColor White
    Write-Host "  [6]  Move Repo             Move a repo to a different collection" -ForegroundColor White
    Write-Host "  [7]  Push to GitHub        Send a converted repo to GitHub" -ForegroundColor White
    Write-Host ""
    Write-Host "  ── Batch (from spreadsheet) ───────────────────────────" -ForegroundColor DarkGray
    Write-Host "  [8]  Run Migration Plan    Execute a saved migration plan file" -ForegroundColor White
    Write-Host "  [9]  Batch Migrate         Migrate all repos from the MDR spreadsheet" -ForegroundColor White
    Write-Host "  [10] Batch Archive         Archive repos from the Dalptfs01 spreadsheet" -ForegroundColor White
    Write-Host ""
    Write-Host "  ── Other ──────────────────────────────────────────────" -ForegroundColor DarkGray
    Write-Host "  [11] View Logs             Open the logs folder" -ForegroundColor White
    Write-Host "  [0]  Exit" -ForegroundColor DarkGray
    Write-Host ""
    Write-Host "  Select an option: " -ForegroundColor Yellow -NoNewline
    $choice = Read-Host

    switch ($choice.Trim()) {
        '1' {
            # Setup Wizard
            & "$PSScriptRoot/New-MigrationConfig.ps1"
        }
        '2' {
            # Prerequisites
            Show-Banner
            & "$PSScriptRoot/Install-Prerequisites.ps1"
            Pause-ForUser
        }
        '3' {
            # Discovery
            if (-not (Test-ConfigReady)) { continue }
            Show-Banner
            & "$PSScriptRoot/Invoke-TfvcDiscovery.ps1" -ConfigPath $ConfigPath -Interactive
            Pause-ForUser
        }
        '4' {
            # Convert
            if (-not (Test-ConfigReady)) { continue }
            Show-Banner
            & "$PSScriptRoot/Convert-TfvcToGit.ps1" -ConfigPath $ConfigPath -Interactive
            Pause-ForUser
        }
        '5' {
            # Split
            if (-not (Test-ConfigReady)) { continue }
            Show-Banner
            & "$PSScriptRoot/Split-TfvcToGitRepos.ps1" -ConfigPath $ConfigPath -Interactive
            Pause-ForUser
        }
        '6' {
            # Move
            if (-not (Test-ConfigReady)) { continue }
            Show-Banner
            & "$PSScriptRoot/Move-RepoToCollection.ps1" -ConfigPath $ConfigPath -Interactive
            Pause-ForUser
        }
        '7' {
            # Push to GitHub
            if (-not (Test-ConfigReady)) { continue }
            Show-Banner
            & "$PSScriptRoot/Push-ToGitHub.ps1" -ConfigPath $ConfigPath -Interactive
            Pause-ForUser
        }
        '8' {
            # Batch migration
            if (-not (Test-ConfigReady)) { continue }
            Show-Banner

            Write-Host "  Enter the path to your migration plan JSON file:" -ForegroundColor Yellow
            Write-Host "  (e.g., ./config/migration-plan.json)" -ForegroundColor DarkGray
            Write-Host ""
            Write-Host "  Path: " -ForegroundColor Yellow -NoNewline
            $planPath = Read-Host

            if (-not $planPath -or -not (Test-Path $planPath)) {
                Write-Host "  File not found: $planPath" -ForegroundColor Red
                Pause-ForUser
                continue
            }

            Write-Host ""
            Write-Host "  Run as dry-run first (preview only, no changes)? [Y/n]: " -ForegroundColor Yellow -NoNewline
            $dryRun = Read-Host

            if ($dryRun.Trim() -match '^[Nn]') {
                & "$PSScriptRoot/Start-Migration.ps1" -ConfigPath $ConfigPath -PlanFile $planPath
            }
            else {
                & "$PSScriptRoot/Start-Migration.ps1" -ConfigPath $ConfigPath -PlanFile $planPath -DryRun
                Write-Host ""
                Write-Host "  That was a preview. Run again without dry-run to execute? [y/N]: " -ForegroundColor Yellow -NoNewline
                $runForReal = Read-Host
                if ($runForReal.Trim() -match '^[Yy]') {
                    & "$PSScriptRoot/Start-Migration.ps1" -ConfigPath $ConfigPath -PlanFile $planPath
                }
            }
            Pause-ForUser
        }
        '9' {
            # Excel-driven migration
            if (-not (Test-ConfigReady)) { continue }
            Show-Banner
            & "$PSScriptRoot/Invoke-ExcelMigration.ps1" -ConfigPath $ConfigPath -Interactive
            Pause-ForUser
        }
        '10' {
            # Archive repos
            if (-not (Test-ConfigReady)) { continue }
            Show-Banner
            & "$PSScriptRoot/Invoke-ArchiveRepos.ps1" -ConfigPath $ConfigPath -Interactive
            Pause-ForUser
        }
        '11' {
            # View logs
            $logDir = './logs'
            if (Test-Path $ConfigPath) {
                try {
                    $cfg = Read-MigrationConfig -ConfigPath $ConfigPath
                    $logDir = $cfg.logDirectory
                }
                catch { }
            }

            if (Test-Path $logDir) {
                $logs = Get-ChildItem $logDir -Filter '*.log' | Sort-Object LastWriteTime -Descending | Select-Object -First 10
                if ($logs.Count -eq 0) {
                    Write-Host "  No log files found." -ForegroundColor Yellow
                }
                else {
                    Write-Host ""
                    Write-Host "  Recent logs:" -ForegroundColor White
                    Write-Host ""
                    for ($i = 0; $i -lt $logs.Count; $i++) {
                        $log = $logs[$i]
                        $size = [math]::Round($log.Length / 1KB, 1)
                        Write-Host "  [$($i + 1)] $($log.Name)  (${size} KB, $($log.LastWriteTime.ToString('MMM dd HH:mm')))" -ForegroundColor White
                    }
                    Write-Host ""
                    Write-Host "  Enter a number to view, or press Enter to go back: " -ForegroundColor Yellow -NoNewline
                    $logChoice = Read-Host

                    if ($logChoice.Trim() -match '^\d+$') {
                        $idx = [int]$logChoice.Trim() - 1
                        if ($idx -ge 0 -and $idx -lt $logs.Count) {
                            Write-Host ""
                            Get-Content $logs[$idx].FullName -Tail 50
                        }
                    }
                }
            }
            else {
                Write-Host "  Log directory not found: $logDir" -ForegroundColor Yellow
            }
            Pause-ForUser
        }
        '0' {
            Write-Host ""
            Write-Host "  Goodbye!" -ForegroundColor Cyan
            Write-Host ""
            return
        }
        default {
            Write-Host "  Invalid option. Please enter a number 0-11." -ForegroundColor Red
            Start-Sleep -Seconds 1
        }
    }
}
