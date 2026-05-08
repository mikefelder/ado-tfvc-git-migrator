#Requires -Version 7.0
<#
.SYNOPSIS
    Splits specific folders within a TFVC repo into separate Git repositories.

.DESCRIPTION
    For repos that contain multiple applications/services in subfolders, this script:
    1. Clones the full TFVC repo using git-tfs
    2. Uses git filter-repo (or filter-branch) to extract each specified subfolder
       into its own standalone Git repo with only the relevant history
    3. Each output repo is independent and ready for GitHub push

    Supports an -Interactive mode that lets you browse TFVC folders, multi-select
    which ones to split out, name each output repo, and optionally choose a
    destination collection/project for each.

.PARAMETER ConfigPath
    Path to migration-config.json.

.PARAMETER Interactive
    Launch interactive mode — browse and select folders via menus.

.PARAMETER Collection
    ADO collection name (skipped in interactive mode).

.PARAMETER ProjectName
    Team project name (skipped in interactive mode).

.PARAMETER TfvcPath
    Root TFVC path to clone (e.g., $/MyProject).

.PARAMETER FolderMappings
    Hashtable mapping TFVC subfolder paths to output repo names.
    Example: @{ '$/MyProject/AppA' = 'app-a'; '$/MyProject/AppB' = 'app-b' }

.PARAMETER MappingsFile
    Alternative to FolderMappings: a JSON file containing the mappings.
    Format: { "$/MyProject/AppA": "app-a", "$/MyProject/AppB": "app-b" }

.PARAMETER HistoryDepth
    Number of changesets to include. Null = full history.

.EXAMPLE
    # Interactive mode — guided menus
    ./Split-TfvcToGitRepos.ps1 -ConfigPath ./config/migration-config.json -Interactive

.EXAMPLE
    # Direct mode — all parameters specified
    ./Split-TfvcToGitRepos.ps1 -ConfigPath ./config/migration-config.json `
        -Collection GAMS -ProjectName "MonoRepo" -TfvcPath "$/MonoRepo" `
        -FolderMappings @{ '$/MonoRepo/ServiceA' = 'service-a'; '$/MonoRepo/ServiceB' = 'service-b' }

.EXAMPLE
    ./Split-TfvcToGitRepos.ps1 -ConfigPath ./config/migration-config.json `
        -Collection GAMS -ProjectName "MonoRepo" -TfvcPath "$/MonoRepo" `
        -MappingsFile ./config/folder-mappings.json
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory)]
    [string]$ConfigPath,

    [switch]$Interactive,

    [string]$Collection,
    [string]$ProjectName,
    [string]$TfvcPath,

    [hashtable]$FolderMappings,

    [string]$MappingsFile,

    [int]$HistoryDepth
)

$ErrorActionPreference = 'Stop'
Import-Module "$PSScriptRoot/modules/AdoTfvcMigrator.psm1" -Force

# ─── Load Config ───────────────────────────────────────────────────────────────

$config = Read-MigrationConfig -ConfigPath $ConfigPath
$logFile = Initialize-MigrationLog -LogDirectory $config.logDirectory -ScriptName 'SplitTfvcToGit'

Write-MigrationLog -Message "Starting TFVC folder split" -LogFile $logFile -Level INFO

# ─── Interactive Mode ──────────────────────────────────────────────────────────

if ($Interactive) {
    Show-MenuHeader -Title "Split TFVC Repo Into Separate Git Repos"
    Write-Host "This wizard will walk you through selecting a TFVC project," -ForegroundColor DarkGray
    Write-Host "picking folders to split out, and naming each output repo." -ForegroundColor DarkGray

    # ── 1. Pick collection ──
    $Collection = Select-AdoCollection -Config $config -Prompt 'Select collection'
    if (-not $Collection) { Write-Host "Cancelled." -ForegroundColor Yellow; return }

    # ── 2. Pick project ──
    $ProjectName = Select-AdoProject -Config $config -Collection $Collection -Prompt 'Select project'
    if (-not $ProjectName) { Write-Host "Cancelled." -ForegroundColor Yellow; return }

    $TfvcPath = "`$/$ProjectName"

    # ── 3. Browse folders and optionally drill down ──
    $currentPath = $TfvcPath
    $selectedFolders = $null

    while ($true) {
        $selectedFolders = Select-TfvcFolders -Config $config -Collection $Collection `
            -ProjectName $ProjectName -ParentPath $currentPath `
            -Prompt 'Select folder(s) to split into separate repos' -MultiSelect
        if (-not $selectedFolders) { Write-Host "Cancelled." -ForegroundColor Yellow; return }

        # Show what was selected and allow drill-down or confirm
        Write-Host ""
        Write-Host "  Selected $($selectedFolders.Count) folder(s):" -ForegroundColor Green
        foreach ($f in $selectedFolders) {
            Write-Host "    • $f" -ForegroundColor White
        }
        Write-Host ""
        Write-Host "  [1] Confirm these selections" -ForegroundColor White
        Write-Host "  [2] Drill into a folder to pick subfolders instead" -ForegroundColor White
        Write-Host "  [3] Start over from project root" -ForegroundColor White
        Write-Host "  [0] Cancel" -ForegroundColor DarkGray
        Write-Host ""
        Write-Host "Choice: " -ForegroundColor Yellow -NoNewline
        $drillChoice = Read-Host

        switch ($drillChoice.Trim()) {
            '1' { break }
            '2' {
                if ($selectedFolders.Count -eq 1) {
                    $currentPath = $selectedFolders[0]
                }
                else {
                    Write-Host ""
                    Write-Host "  Pick which folder to drill into:" -ForegroundColor Yellow
                    $drillIdx = Show-NumberedMenu -Items $selectedFolders -Prompt 'Drill into' -AllowBack
                    if ($null -ne $drillIdx) {
                        $currentPath = $selectedFolders[$drillIdx]
                    }
                }
                $selectedFolders = $null
                continue
            }
            '3' {
                $currentPath = $TfvcPath
                $selectedFolders = $null
                continue
            }
            '0' { Write-Host "Cancelled." -ForegroundColor Yellow; return }
            default { break }
        }
        if ($selectedFolders) { break }
    }

    # ── 4. Name each output repo ──
    $FolderMappings = @{}

    Show-MenuHeader -Title "Name Output Repositories"
    Write-Host "  For each selected folder, provide a Git repo name." -ForegroundColor DarkGray
    Write-Host "  Press Enter to accept the suggested default." -ForegroundColor DarkGray
    Write-Host ""

    foreach ($folder in $selectedFolders) {
        $leafName = ($folder -split '/')[-1].ToLower()
        $defaultName = ($leafName -replace '[^a-z0-9\-]', '-')

        Write-Host "  $folder" -ForegroundColor White
        Write-Host "    Repo name [$defaultName]: " -ForegroundColor Yellow -NoNewline
        $nameInput = Read-Host
        $repoName = if ($nameInput.Trim()) { $nameInput.Trim() } else { $defaultName }
        $FolderMappings[$folder] = $repoName
        Write-Host "    → $repoName" -ForegroundColor Green
        Write-Host ""
    }

    # ── 5. History depth ──
    Write-Host "  History depth (enter for full history, or a number): " -ForegroundColor Yellow -NoNewline
    $depthInput = Read-Host
    if ($depthInput.Trim() -match '^\d+$') {
        $HistoryDepth = [int]$depthInput.Trim()
    }

    # ── 6. Optional: choose destination collection/project for each ──
    Write-Host ""
    Write-Host "  Push split repos to a different ADO collection after splitting? [y/N]: " -ForegroundColor Yellow -NoNewline
    $pushToAdo = Read-Host

    $adoDestinations = @{}
    if ($pushToAdo.Trim() -match '^[Yy]') {
        $destCollection = Select-AdoCollection -Config $config -Prompt 'Select DESTINATION collection'
        if ($destCollection) {
            $destProject = Select-AdoProject -Config $config -Collection $destCollection -Prompt 'Select DESTINATION project'
            if ($destProject) {
                foreach ($folder in $selectedFolders) {
                    $adoDestinations[$folder] = @{
                        Collection = $destCollection
                        Project    = $destProject
                        RepoName   = $FolderMappings[$folder]
                    }
                }
            }
        }
    }

    # ── Confirm ──
    Show-MenuHeader -Title "Confirm Split Plan"
    Write-Host "  Source: $Collection / $ProjectName ($TfvcPath)" -ForegroundColor White
    if ($HistoryDepth) {
        Write-Host "  History: $HistoryDepth changesets" -ForegroundColor White
    }
    else {
        Write-Host "  History: Full" -ForegroundColor White
    }
    Write-Host ""
    Write-Host "  Splits:" -ForegroundColor White
    foreach ($folder in $FolderMappings.Keys) {
        $line = "    $folder → $($FolderMappings[$folder])"
        if ($adoDestinations.ContainsKey($folder)) {
            $dest = $adoDestinations[$folder]
            $line += " → push to $($dest.Collection)/$($dest.Project)"
        }
        Write-Host $line -ForegroundColor Cyan
    }
    Write-Host ""
    Write-Host "  Proceed? [Y/n]: " -ForegroundColor Yellow -NoNewline
    $confirm = Read-Host
    if ($confirm.Trim() -match '^[Nn]') {
        Write-Host "Cancelled." -ForegroundColor Yellow
        return
    }
}

# ─── Validate required params in non-interactive mode ──────────────────────────

if (-not $Collection)  { throw "Collection is required. Use -Interactive or provide -Collection." }
if (-not $ProjectName) { throw "ProjectName is required. Use -Interactive or provide -ProjectName." }
if (-not $TfvcPath)    { throw "TfvcPath is required. Use -Interactive or provide -TfvcPath." }

# ─── Resolve Mappings ─────────────────────────────────────────────────────────

if (-not $FolderMappings -and -not $MappingsFile) {
    throw "Provide -FolderMappings, -MappingsFile, or use -Interactive."
}

if ($MappingsFile) {
    if (-not (Test-Path $MappingsFile)) {
        throw "Mappings file not found: $MappingsFile"
    }
    $FolderMappings = Get-Content $MappingsFile -Raw | ConvertFrom-Json -AsHashtable
}

Write-MigrationLog -Message "Folder mappings:" -LogFile $logFile
foreach ($key in $FolderMappings.Keys) {
    Write-MigrationLog -Message "  $key → $($FolderMappings[$key])" -LogFile $logFile
}

# ─── Validate ──────────────────────────────────────────────────────────────────

$collectionConfig = $config.collections[$Collection]
if (-not $collectionConfig) {
    throw "Collection '$Collection' not found in config."
}

$pat = $collectionConfig.pat
$gitTfsPath = Find-GitTfs -GitTfsPath $config.gitTfsPath

# Check for git-filter-repo
$hasFilterRepo = $null -ne (Get-Command 'git-filter-repo' -ErrorAction SilentlyContinue)
if (-not $hasFilterRepo) {
    Write-MigrationLog -Message "git-filter-repo not found. Will use git filter-branch (slower). Install: pip install git-filter-repo" -LogFile $logFile -Level WARN
}

# ─── Step 1: Clone the full TFVC repo ─────────────────────────────────────────

$stagingName = "staging-$(($TfvcPath -replace '^\$/', '' -replace '/', '-').ToLower())"
$stagingPath = Join-Path $config.outputDirectory $stagingName

Write-MigrationLog -Message "Step 1: Cloning full TFVC repo to staging area" -LogFile $logFile -Level INFO

if (Test-Path $stagingPath) {
    Write-MigrationLog -Message "Staging path exists: $stagingPath" -LogFile $logFile -Level WARN
    Write-Host "  The staging directory already exists: $stagingPath" -ForegroundColor Yellow
    Write-Host "  Delete it and start fresh? [Y/n]: " -ForegroundColor Yellow -NoNewline
    $confirm = Read-Host
    if ($confirm.Trim() -match '^[Nn]') {
        Write-Host "  Cancelled." -ForegroundColor Yellow
        return
    }
    Remove-Item -Recurse -Force $stagingPath
}

$tfsUrl = "$($config.adoServerUrl)/$Collection"
$cloneArgs = "clone `"$tfsUrl`" `"$TfvcPath`" `"$stagingPath`""

if ($HistoryDepth) {
    $cloneArgs += " --changeset=$HistoryDepth"
}

$authorFile = $config.authorMappingFile
if ($authorFile -and (Test-Path $authorFile)) {
    $cloneArgs += " --authors=`"$authorFile`""
}

# Authenticate with PAT
$env:GIT_TFS_USERNAME = ''
$env:GIT_TFS_PASSWORD = $pat
$cloneArgs += " --username=`"`" --password=`"$pat`""

try {
    Invoke-GitTfs -Arguments $cloneArgs -LogFile $logFile
    Write-MigrationLog -Message "Staging clone complete" -LogFile $logFile -Level SUCCESS
}
catch {
    Write-MigrationLog -Message "git-tfs clone failed: $($_.Exception.Message)" -LogFile $logFile -Level ERROR
    throw
}
finally {
    Remove-Item Env:\GIT_TFS_USERNAME -ErrorAction SilentlyContinue
    Remove-Item Env:\GIT_TFS_PASSWORD -ErrorAction SilentlyContinue
}

# Clean git-tfs metadata from staging
Remove-GitTfsMetadata -RepoPath $stagingPath -LogFile $logFile

# ─── Step 2: Split each folder into its own repo ──────────────────────────────

$defaultBranch = $config.defaultBranch ?? 'main'
$results = [System.Collections.ArrayList]::new()

foreach ($tfvcFolder in $FolderMappings.Keys) {
    $repoName = $FolderMappings[$tfvcFolder]
    $outputPath = Join-Path $config.outputDirectory $repoName

    Write-MigrationLog -Message "Step 2: Extracting '$tfvcFolder' → '$repoName'" -LogFile $logFile -Level INFO

    if (Test-Path $outputPath) {
        Write-Host "  Output directory already exists: $outputPath" -ForegroundColor Yellow
        Write-Host "  Overwrite? [Y/n]: " -ForegroundColor Yellow -NoNewline
        $overwrite = Read-Host
        if ($overwrite.Trim() -match '^[Nn]') {
            Write-MigrationLog -Message "  Skipping '$repoName' — output directory already exists" -LogFile $logFile -Level WARN
            continue
        }
        Remove-Item -Recurse -Force $outputPath
    }

    # Copy the staging repo
    Copy-Item -Path $stagingPath -Destination $outputPath -Recurse

    # Determine the relative subfolder path within the repo
    # TFVC path: $/Project/FolderA → relative path in git clone is typically FolderA
    $relativePath = $tfvcFolder -replace [regex]::Escape($TfvcPath), '' -replace '^/', ''

    if (-not $relativePath) {
        Write-MigrationLog -Message "  No subfolder extraction needed (root path)" -LogFile $logFile
    }
    elseif ($hasFilterRepo) {
        # Use git-filter-repo (fast, recommended)
        Write-MigrationLog -Message "  Using git-filter-repo for path: $relativePath" -LogFile $logFile
        try {
            $filterArgs = "filter-repo --subdirectory-filter `"$relativePath`" --force"
            Invoke-Git -Arguments $filterArgs -WorkingDirectory $outputPath -LogFile $logFile
        }
        catch {
            Write-MigrationLog -Message "  git-filter-repo failed: $($_.Exception.Message)" -LogFile $logFile -Level ERROR
            Write-MigrationLog -Message "  Falling back to git filter-branch" -LogFile $logFile -Level WARN
            # Fallback
            $filterArgs = "filter-branch --subdirectory-filter `"$relativePath`" --prune-empty -- --all"
            Invoke-Git -Arguments $filterArgs -WorkingDirectory $outputPath -LogFile $logFile
        }
    }
    else {
        # Use git filter-branch (slower fallback)
        Write-MigrationLog -Message "  Using git filter-branch for path: $relativePath" -LogFile $logFile
        $filterArgs = "filter-branch --subdirectory-filter `"$relativePath`" --prune-empty -- --all"
        Invoke-Git -Arguments $filterArgs -WorkingDirectory $outputPath -LogFile $logFile
    }

    # Rename branch
    try {
        Invoke-Git -Arguments "branch -m master $defaultBranch" -WorkingDirectory $outputPath -LogFile $logFile
    }
    catch { }

    # Garbage collect
    Invoke-Git -Arguments "reflog expire --expire=now --all" -WorkingDirectory $outputPath -LogFile $logFile
    Invoke-Git -Arguments "gc --prune=now --aggressive" -WorkingDirectory $outputPath -LogFile $logFile

    $commitCount = (Invoke-Git -Arguments "rev-list --count HEAD" -WorkingDirectory $outputPath -LogFile $logFile).Trim()

    [void]$results.Add([PSCustomObject]@{
        TfvcPath   = $tfvcFolder
        GitRepo    = $repoName
        OutputPath = $outputPath
        Commits    = $commitCount
        Status     = 'Success'
    })

    Write-MigrationLog -Message "  Done: $repoName ($commitCount commits)" -LogFile $logFile -Level SUCCESS
}

# ─── Cleanup staging ──────────────────────────────────────────────────────────

Write-MigrationLog -Message "Removing staging directory: $stagingPath" -LogFile $logFile
Remove-Item -Recurse -Force $stagingPath

# ─── Summary ──────────────────────────────────────────────────────────────────

Write-Host ""
Write-MigrationLog -Message "Split complete! Results:" -LogFile $logFile -Level SUCCESS
$results | Format-Table -AutoSize

# ─── Push to ADO destinations (interactive mode) ──────────────────────────────

if ($Interactive -and $adoDestinations -and $adoDestinations.Count -gt 0) {
    Write-Host ""
    Write-MigrationLog -Message "Pushing split repos to destination collection..." -LogFile $logFile -Level INFO

    foreach ($folder in $adoDestinations.Keys) {
        $dest = $adoDestinations[$folder]
        $repoName = $dest.RepoName
        $repoPath = Join-Path $config.outputDirectory $repoName

        if (Test-Path (Join-Path $repoPath '.git')) {
            try {
                & "$PSScriptRoot/Move-RepoToCollection.ps1" `
                    -ConfigPath $ConfigPath `
                    -SourceCollection $Collection `
                    -SourceProject $ProjectName `
                    -TfvcPath $folder `
                    -TargetCollection $dest.Collection `
                    -TargetProject $dest.Project `
                    -TargetRepoName $repoName `
                    -SkipTargetRepoCreation:$false

                Write-MigrationLog -Message "Pushed $repoName to $($dest.Collection)/$($dest.Project)" -LogFile $logFile -Level SUCCESS
            }
            catch {
                Write-MigrationLog -Message "Failed to push $repoName — $($_.Exception.Message)" -LogFile $logFile -Level ERROR
            }
        }
    }
}

Write-Host ""
Write-Host "Next steps — push each repo to GitHub:" -ForegroundColor Cyan
foreach ($result in $results) {
    Write-Host "  ./Push-ToGitHub.ps1 -RepoPath `"$($result.OutputPath)`" -GitHubOrg McDermott -GitHubRepo $($result.GitRepo)" -ForegroundColor Cyan
}

Write-MigrationLog -Message "Log file: $logFile" -LogFile $logFile -Level INFO
