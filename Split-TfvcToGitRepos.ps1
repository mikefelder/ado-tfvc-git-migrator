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

.PARAMETER ConfigPath
    Path to migration-config.json.

.PARAMETER Collection
    ADO collection name.

.PARAMETER ProjectName
    Team project name.

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

    [Parameter(Mandatory)]
    [string]$Collection,

    [Parameter(Mandatory)]
    [string]$ProjectName,

    [Parameter(Mandatory)]
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

# ─── Resolve Mappings ─────────────────────────────────────────────────────────

if (-not $FolderMappings -and -not $MappingsFile) {
    throw "Provide either -FolderMappings or -MappingsFile."
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
    Write-MigrationLog -Message "Staging path exists, removing: $stagingPath" -LogFile $logFile -Level WARN
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

$env:GIT_TFS_PAT = $pat

try {
    Invoke-GitTfs -Arguments $cloneArgs -LogFile $logFile
    Write-MigrationLog -Message "Staging clone complete" -LogFile $logFile -Level SUCCESS
}
catch {
    Write-MigrationLog -Message "git-tfs clone failed: $($_.Exception.Message)" -LogFile $logFile -Level ERROR
    throw
}
finally {
    Remove-Item Env:\GIT_TFS_PAT -ErrorAction SilentlyContinue
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

Write-Host ""
Write-Host "Next steps — push each repo to GitHub:" -ForegroundColor Cyan
foreach ($result in $results) {
    Write-Host "  ./Push-ToGitHub.ps1 -RepoPath `"$($result.OutputPath)`" -GitHubOrg McDermott -GitHubRepo $($result.GitRepo)" -ForegroundColor Cyan
}

Write-MigrationLog -Message "Log file: $logFile" -LogFile $logFile -Level INFO
