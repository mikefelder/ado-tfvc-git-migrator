#Requires -Version 7.0
<#
.SYNOPSIS
    Interactive wizard to create or update the migration-config.json file.

.DESCRIPTION
    Walks through each section of the configuration file, asking for values
    in plain English. Creates a valid migration-config.json in the config/ folder.

.PARAMETER ConfigPath
    Path where the config file will be saved. Default: ./config/migration-config.json

.EXAMPLE
    ./New-MigrationConfig.ps1
#>

[CmdletBinding()]
param(
    [string]$ConfigPath = (Join-Path $PSScriptRoot 'config/migration-config.json')
)

$ErrorActionPreference = 'Stop'
Import-Module "$PSScriptRoot/modules/AdoTfvcMigrator.psm1" -Force

# ─── Helpers ───────────────────────────────────────────────────────────────────

function Read-WithDefault {
    param(
        [string]$Prompt,
        [string]$Default
    )
    if ($Default) {
        Write-Host "  ${Prompt} [$Default]: " -ForegroundColor Yellow -NoNewline
    }
    else {
        Write-Host "  ${Prompt}: " -ForegroundColor Yellow -NoNewline
    }
    $val = Read-Host
    if ([string]::IsNullOrWhiteSpace($val)) { return $Default }
    return $val.Trim()
}

function Read-SecureValue {
    param(
        [string]$Prompt,
        [string]$Current
    )
    if ($Current -and $Current -ne 'YOUR_PAT_HERE') {
        $masked = $Current.Substring(0, [Math]::Min(4, $Current.Length)) + ('*' * [Math]::Max(0, $Current.Length - 4))
        Write-Host "  ${Prompt} (current: $masked)" -ForegroundColor Yellow
        Write-Host "  Press Enter to keep current, or paste a new value: " -ForegroundColor Yellow -NoNewline
    }
    else {
        Write-Host "  ${Prompt}: " -ForegroundColor Yellow -NoNewline
    }
    $val = Read-Host
    if ([string]::IsNullOrWhiteSpace($val)) { return $Current }
    return $val.Trim()
}

# ─── Load existing config or start fresh ──────────────────────────────────────

$config = $null
$isUpdate = $false

if (Test-Path $ConfigPath) {
    Write-Host ""
    Write-Host "  An existing config file was found:" -ForegroundColor White
    Write-Host "  $ConfigPath" -ForegroundColor DarkGray
    Write-Host ""
    Write-Host "  [1] Update the existing configuration" -ForegroundColor White
    Write-Host "  [2] Start fresh (existing file will be backed up)" -ForegroundColor White
    Write-Host "  [0] Cancel" -ForegroundColor DarkGray
    Write-Host ""
    Write-Host "  Choice: " -ForegroundColor Yellow -NoNewline
    $updateChoice = Read-Host

    switch ($updateChoice.Trim()) {
        '0' { Write-Host "  Cancelled." -ForegroundColor Yellow; return }
        '2' {
            $backupPath = "$ConfigPath.backup.$(Get-Date -Format 'yyyyMMdd-HHmmss')"
            Copy-Item -Path $ConfigPath -Destination $backupPath
            Write-Host "  Backed up to: $backupPath" -ForegroundColor DarkGray
        }
        default {
            try {
                $raw = Get-Content $ConfigPath -Raw | ConvertFrom-Json -AsHashtable
                $config = $raw
                $isUpdate = $true
                Write-Host "  Loaded existing config." -ForegroundColor Green
            }
            catch {
                Write-Host "  Could not parse existing config — starting fresh." -ForegroundColor Yellow
            }
        }
    }
}

# ─── Initialize config structure ──────────────────────────────────────────────

if (-not $config) {
    $config = @{
        adoServerUrl      = ''
        collections       = @{}
        outputDirectory   = './output'
        logDirectory      = './logs'
        gitTfsPath        = $null
        authorMappingFile = $null
        defaultBranch     = 'main'
        github            = @{
            enterpriseUrl = ''
            apiUrl        = ''
            pat           = ''
            defaultOrg    = ''
        }
        migrationDefaults = @{
            includeHistory    = $true
            historyDepthLimit = $null
            excludePatterns   = @('**/bin/**', '**/obj/**', '**/packages/**', '**/.vs/**', '**/node_modules/**')
            gitLfsExtensions  = @('*.dll', '*.exe', '*.zip', '*.nupkg', '*.pdb')
            cleanupBinaries   = $true
        }
    }
}

# ─── Step 1: ADO Server ───────────────────────────────────────────────────────

Show-MenuHeader -Title "Step 1 of 4: ADO Server Connection"

Write-Host "  This is the base URL of your Azure DevOps Server 2022 instance." -ForegroundColor DarkGray
Write-Host "  Example: https://ado.mcdermott.com" -ForegroundColor DarkGray
Write-Host ""

$config.adoServerUrl = Read-WithDefault -Prompt 'ADO Server URL' -Default $config.adoServerUrl

if ([string]::IsNullOrWhiteSpace($config.adoServerUrl)) {
    Write-Host "  ADO Server URL is required. Cannot continue." -ForegroundColor Red
    return
}

# ─── Step 2: Collections ──────────────────────────────────────────────────────

Show-MenuHeader -Title "Step 2 of 4: ADO Collections"

Write-Host "  Each ADO collection needs a name and a Personal Access Token (PAT)." -ForegroundColor DarkGray
Write-Host "  PATs can be created in ADO under your profile → Security → Personal Access Tokens." -ForegroundColor DarkGray
Write-Host "  Required scopes: Code (Read & Write), Project and Team (Read)." -ForegroundColor DarkGray
Write-Host ""

if ($config.collections.Count -gt 0) {
    Write-Host "  Current collections:" -ForegroundColor White
    foreach ($name in $config.collections.Keys | Sort-Object) {
        $desc = $config.collections[$name].description
        $patStatus = if ($config.collections[$name].pat -and $config.collections[$name].pat -ne 'YOUR_PAT_HERE') { 'PAT set' } else { 'PAT needed' }
        Write-Host "    • $name ($patStatus)" -ForegroundColor DarkGray
    }
    Write-Host ""
}

$addMore = $true
while ($addMore) {
    $collName = Read-WithDefault -Prompt 'Collection name (or press Enter when done)'
    if ([string]::IsNullOrWhiteSpace($collName)) {
        if ($config.collections.Count -eq 0) {
            Write-Host "  You need at least one collection. Please enter a name." -ForegroundColor Red
            continue
        }
        $addMore = $false
        continue
    }

    $existing = $config.collections[$collName]
    $existingPat = if ($existing) { $existing.pat } else { '' }
    $existingDesc = if ($existing) { $existing.description } else { '' }

    $collPat = Read-SecureValue -Prompt "PAT for '$collName'" -Current $existingPat
    $collDesc = Read-WithDefault -Prompt "Description for '$collName' (optional)" -Default $existingDesc

    $config.collections[$collName] = @{
        pat         = $collPat
        description = $collDesc
    }

    Write-Host "  ✓ Collection '$collName' configured." -ForegroundColor Green
    Write-Host ""

    # Test connection
    Write-Host "  Testing connection to '$collName'... " -ForegroundColor DarkGray -NoNewline
    if ($collPat -and $collPat -ne 'YOUR_PAT_HERE') {
        $testResult = Test-AdoConnection -ServerUrl $config.adoServerUrl -Collection $collName -Pat $collPat
        if ($testResult.Connected) {
            Write-Host "Connected! ($($testResult.ProjectCount) projects found)" -ForegroundColor Green
        }
        else {
            Write-Host "Failed: $($testResult.Error)" -ForegroundColor Red
            Write-Host "  You can continue setup and fix the PAT later." -ForegroundColor Yellow
        }
    }
    else {
        Write-Host "Skipped (no PAT provided)" -ForegroundColor Yellow
    }
    Write-Host ""
}

# ─── Step 3: GitHub Enterprise ────────────────────────────────────────────────

Show-MenuHeader -Title "Step 3 of 4: GitHub Enterprise (Optional)"

Write-Host "  If you're pushing repos to GitHub Enterprise, configure it here." -ForegroundColor DarkGray
Write-Host "  Skip this section if you're only converting repos locally." -ForegroundColor DarkGray
Write-Host ""
Write-Host "  Configure GitHub Enterprise? [y/N]: " -ForegroundColor Yellow -NoNewline
$ghChoice = Read-Host

if ($ghChoice.Trim() -match '^[Yy]') {
    $ghUrl = Read-WithDefault -Prompt 'GitHub Enterprise URL (e.g. https://github.mcdermott.com)' -Default $config.github.enterpriseUrl
    $config.github.enterpriseUrl = $ghUrl

    # Auto-derive API URL
    $defaultApiUrl = if ($ghUrl) { "$ghUrl/api/v3" } else { '' }
    if (-not $config.github.apiUrl -or $config.github.apiUrl -eq '') {
        $config.github.apiUrl = $defaultApiUrl
    }
    $config.github.apiUrl = Read-WithDefault -Prompt 'GitHub API URL' -Default $config.github.apiUrl

    $config.github.pat = Read-SecureValue -Prompt 'GitHub PAT (needs repo scope)' -Current $config.github.pat
    $config.github.defaultOrg = Read-WithDefault -Prompt 'Default GitHub organization' -Default $config.github.defaultOrg
}

# ─── Step 4: Directories & Defaults ──────────────────────────────────────────

Show-MenuHeader -Title "Step 4 of 4: Output Settings"

Write-Host "  These defaults work for most migrations. Press Enter to accept each default." -ForegroundColor DarkGray
Write-Host ""

$config.outputDirectory = Read-WithDefault -Prompt 'Output directory (where converted repos go)' -Default $config.outputDirectory
$config.logDirectory = Read-WithDefault -Prompt 'Log directory' -Default $config.logDirectory
$config.defaultBranch = Read-WithDefault -Prompt 'Default Git branch name' -Default $config.defaultBranch

# ─── Save ─────────────────────────────────────────────────────────────────────

Write-Host ""
Write-Host "  ─── Configuration Summary ─────────────────────────────────" -ForegroundColor Cyan
Write-Host ""
Write-Host "  ADO Server:     $($config.adoServerUrl)" -ForegroundColor White
Write-Host "  Collections:    $($config.collections.Keys -join ', ')" -ForegroundColor White
if ($config.github.enterpriseUrl) {
    Write-Host "  GitHub:         $($config.github.enterpriseUrl) (org: $($config.github.defaultOrg))" -ForegroundColor White
}
else {
    Write-Host "  GitHub:         Not configured" -ForegroundColor DarkGray
}
Write-Host "  Output dir:     $($config.outputDirectory)" -ForegroundColor White
Write-Host "  Log dir:        $($config.logDirectory)" -ForegroundColor White
Write-Host "  Default branch: $($config.defaultBranch)" -ForegroundColor White
Write-Host ""

Write-Host "  Save this configuration? [Y/n]: " -ForegroundColor Yellow -NoNewline
$saveChoice = Read-Host

if ($saveChoice.Trim() -match '^[Nn]') {
    Write-Host "  Configuration was NOT saved." -ForegroundColor Yellow
    return
}

# Ensure config directory exists
$configDir = Split-Path $ConfigPath -Parent
if (-not (Test-Path $configDir)) {
    New-Item -ItemType Directory -Path $configDir -Force | Out-Null
}

$config | ConvertTo-Json -Depth 5 | Set-Content -Path $ConfigPath -Encoding utf8

Write-Host ""
Write-Host "  ✓ Configuration saved to: $ConfigPath" -ForegroundColor Green
Write-Host "  You can now run migrations from the main menu." -ForegroundColor DarkGray
Write-Host ""
