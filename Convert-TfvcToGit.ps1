#Requires -Version 7.0
<#
.SYNOPSIS
    Converts an entire TFVC repository (or branch path) to a Git repository using git-tfs.

.DESCRIPTION
    Uses git-tfs clone to convert a TFVC path to a full Git repo with history.
    Supports:
    - Interactive mode for guided selection
    - Full history or depth-limited conversion
    - Author mapping (DOMAIN\user → Git Name <email>)
    - Post-conversion cleanup (git-tfs metadata removal)
    - Git LFS setup for binary file types

.PARAMETER ConfigPath
    Path to migration-config.json.

.PARAMETER Interactive
    Launch interactive mode — browse and select repos via menus.

.PARAMETER Collection
    ADO collection name (skipped in interactive mode).

.PARAMETER ProjectName
    Team project name (skipped in interactive mode).

.PARAMETER TfvcPath
    TFVC path to convert (e.g., $/MyProject or $/MyProject/Main).

.PARAMETER OutputRepoName
    Name for the output Git repository. Default: derived from TfvcPath.

.PARAMETER HistoryDepth
    Number of changesets to include. Null = full history.

.PARAMETER NoCleanup
    Skip post-conversion cleanup (keep git-tfs metadata).

.EXAMPLE
    # Interactive mode
    ./Convert-TfvcToGit.ps1 -ConfigPath ./config/migration-config.json -Interactive

.EXAMPLE
    ./Convert-TfvcToGit.ps1 -ConfigPath ./config/migration-config.json `
        -Collection GAMS -ProjectName "LegacyApp" -TfvcPath "$/LegacyApp"
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory)]
    [string]$ConfigPath,

    [switch]$Interactive,

    [string]$Collection,
    [string]$ProjectName,
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

# ─── Interactive Mode ──────────────────────────────────────────────────────────

if ($Interactive) {
    Show-MenuHeader -Title "Convert TFVC Repository to Git"
    Write-Host "This wizard will help you select a TFVC repo and convert it to Git." -ForegroundColor DarkGray
    Write-Host "The converted repo will be saved locally, ready to push to GitHub." -ForegroundColor DarkGray

    # 1. Pick collection
    $Collection = Select-AdoCollection -Config $config -Prompt 'Select collection'
    if (-not $Collection) { Write-Host "Cancelled." -ForegroundColor Yellow; return }

    # 2. Pick project
    $ProjectName = Select-AdoProject -Config $config -Collection $Collection -Prompt 'Select project'
    if (-not $ProjectName) { Write-Host "Cancelled." -ForegroundColor Yellow; return }

    # 3. Browse TFVC folders (or use project root)
    $TfvcPath = "`$/$ProjectName"
    Write-Host ""
    Write-Host "  Convert the entire project, or pick a specific folder?" -ForegroundColor White
    Write-Host ""
    Write-Host "  [1] Convert entire project ($/  $ProjectName)" -ForegroundColor White
    Write-Host "  [2] Browse and pick a specific folder" -ForegroundColor White
    Write-Host "  [0] Cancel" -ForegroundColor DarkGray
    Write-Host ""
    Write-Host "  Choice: " -ForegroundColor Yellow -NoNewline
    $scopeChoice = Read-Host

    if ($scopeChoice.Trim() -eq '0') { Write-Host "Cancelled." -ForegroundColor Yellow; return }
    if ($scopeChoice.Trim() -eq '2') {
        $currentPath = $TfvcPath
        while ($true) {
            $selected = Select-TfvcFolders -Config $config -Collection $Collection `
                -ProjectName $ProjectName -ParentPath $currentPath -Prompt 'Select folder to convert'
            if (-not $selected) { Write-Host "Cancelled." -ForegroundColor Yellow; return }

            $chosenPath = $selected[0]
            Write-Host ""
            Write-Host "  Selected: $chosenPath" -ForegroundColor Green
            Write-Host ""
            Write-Host "  [1] Use this folder" -ForegroundColor White
            Write-Host "  [2] Look inside this folder" -ForegroundColor White
            Write-Host "  [0] Cancel" -ForegroundColor DarkGray
            Write-Host ""
            Write-Host "  Choice: " -ForegroundColor Yellow -NoNewline
            $drillChoice = Read-Host

            switch ($drillChoice.Trim()) {
                '1' { $TfvcPath = $chosenPath; break }
                '2' { $currentPath = $chosenPath; continue }
                '0' { Write-Host "Cancelled." -ForegroundColor Yellow; return }
                default { $TfvcPath = $chosenPath; break }
            }
            if ($TfvcPath -ne "`$/$ProjectName") { break }
        }
    }

    # 4. Output repo name
    $defaultName = ($TfvcPath -replace '^\$/', '' -replace '/', '-').ToLower()
    Write-Host ""
    Write-Host "  Git repo name [$defaultName]: " -ForegroundColor Yellow -NoNewline
    $nameInput = Read-Host
    $OutputRepoName = if ($nameInput.Trim()) { $nameInput.Trim() } else { $defaultName }

    # 5. History depth
    Write-Host ""
    Write-Host "  How much history to include?" -ForegroundColor White
    Write-Host "  Press Enter for full history, or type a number (e.g., 500 for last 500 changes)." -ForegroundColor DarkGray
    Write-Host ""
    Write-Host "  History depth: " -ForegroundColor Yellow -NoNewline
    $depthInput = Read-Host
    if ($depthInput.Trim() -match '^\d+$') {
        $HistoryDepth = [int]$depthInput.Trim()
    }

    # 6. Confirm
    Show-MenuHeader -Title "Confirm Conversion"
    Write-Host "  Collection:   $Collection" -ForegroundColor White
    Write-Host "  Project:      $ProjectName" -ForegroundColor White
    Write-Host "  TFVC Path:    $TfvcPath" -ForegroundColor White
    Write-Host "  Output Name:  $OutputRepoName" -ForegroundColor White
    if ($HistoryDepth) {
        Write-Host "  History:      Last $HistoryDepth changes" -ForegroundColor White
    }
    else {
        Write-Host "  History:      Full (all changes)" -ForegroundColor White
    }
    Write-Host ""
    Write-Host "  This will convert the TFVC repo to a local Git repository." -ForegroundColor DarkGray
    Write-Host "  No changes are made to the original TFVC repo." -ForegroundColor DarkGray
    Write-Host ""
    Write-Host "  Proceed? [Y/n]: " -ForegroundColor Yellow -NoNewline
    $confirm = Read-Host
    if ($confirm.Trim() -match '^[Nn]') {
        Write-Host "Cancelled." -ForegroundColor Yellow
        return
    }
}

# ─── Validate required params ─────────────────────────────────────────────────

if (-not $Collection)  { throw "Collection is required. Use -Interactive or provide -Collection." }
if (-not $ProjectName) { throw "ProjectName is required. Use -Interactive or provide -ProjectName." }
if (-not $TfvcPath)    { throw "TfvcPath is required. Use -Interactive or provide -TfvcPath." }

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
Write-Host "  Verifying connection to ADO..." -ForegroundColor DarkGray
$connTest = Test-AdoConnection -ServerUrl $config.adoServerUrl -Collection $Collection -Pat $pat
if (-not $connTest.Connected) {
    Write-Host ""
    Write-Host "  Could not connect to '$Collection'." -ForegroundColor Red
    Write-Host "  This usually means:" -ForegroundColor Yellow
    Write-Host "    • The server URL in your config is wrong" -ForegroundColor Yellow
    Write-Host "    • Your PAT (Personal Access Token) has expired" -ForegroundColor Yellow
    Write-Host "    • The ADO server is unreachable from this machine" -ForegroundColor Yellow
    Write-Host ""
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
    Write-Host ""
    Write-Host "  A previous conversion already exists at: $outputPath" -ForegroundColor Yellow
    Write-Host "  To start fresh, the existing folder needs to be removed." -ForegroundColor Yellow
    Write-Host ""
    Write-Host "  Remove it and continue? [Y/n]: " -ForegroundColor Yellow -NoNewline
    $removeConfirm = Read-Host
    if ($removeConfirm.Trim() -match '^[Nn]') {
        Write-Host "  Cancelled. Existing output left in place." -ForegroundColor Yellow
        return
    }
    Remove-Item -Recurse -Force $outputPath
    Write-MigrationLog -Message "Removed existing directory" -LogFile $logFile
}

# ─── Build git-tfs clone command ───────────────────────────────────────────────

# Enable long paths to handle deeply nested TFVC directories (Windows 248-char limit)
try { & git config --global core.longpaths true 2>$null } catch { }

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

# Authenticate with PAT — set env vars (git-tfs reads these) and pass CLI args as fallback
$env:GIT_TFS_USERNAME = ''
$env:GIT_TFS_PASSWORD = $pat
$cloneArgs += " --username=`"`" --password=`"$pat`""

Write-Host "" -ForegroundColor White
Write-Host "  ┌─────────────────────────────────────────────────────────┐" -ForegroundColor Cyan
Write-Host "  │  Converting TFVC to Git — this may take a while...     │" -ForegroundColor Cyan
Write-Host "  │                                                         │" -ForegroundColor Cyan
Write-Host "  │  • Large repos can take 30+ minutes                    │" -ForegroundColor Cyan
Write-Host "  │  • The screen may appear idle — that's normal           │" -ForegroundColor Cyan
Write-Host "  │  • Progress is being written to the log file            │" -ForegroundColor Cyan
Write-Host "  │  • Do NOT close this window                             │" -ForegroundColor Cyan
Write-Host "  └─────────────────────────────────────────────────────────┘" -ForegroundColor Cyan
Write-Host "" -ForegroundColor White

Write-MigrationLog -Message "Starting git-tfs clone..." -LogFile $logFile -Level INFO
$cloneStart = Get-Date

try {
    Invoke-GitTfs -Arguments $cloneArgs -LogFile $logFile
    $cloneDuration = (Get-Date) - $cloneStart
    Write-MigrationLog -Message "git-tfs clone completed in $($cloneDuration.ToString('hh\:mm\:ss'))" -LogFile $logFile -Level SUCCESS
}
catch {
    Write-Host "" -ForegroundColor Red
    Write-Host "  The conversion failed. Common causes:" -ForegroundColor Red
    Write-Host "    • Network connectivity issue to the ADO server" -ForegroundColor Yellow
    Write-Host "    • The TFVC path doesn't exist or you lack permissions" -ForegroundColor Yellow
    Write-Host "    • git-tfs encountered an unsupported TFVC structure" -ForegroundColor Yellow
    Write-Host "  Check the log file for details: $logFile" -ForegroundColor Yellow
    Write-Host "" -ForegroundColor Red
    Write-MigrationLog -Message "git-tfs clone failed: $($_.Exception.Message)" -LogFile $logFile -Level ERROR
    throw
}
finally {
    Remove-Item Env:\GIT_TFS_USERNAME -ErrorAction SilentlyContinue
    Remove-Item Env:\GIT_TFS_PASSWORD -ErrorAction SilentlyContinue
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
