#Requires -Version 7.0
<#
.SYNOPSIS
    Moves or clones a TFVC project, folder, or Git repo from one ADO collection to another as Git.

.DESCRIPTION
    Handles the scenario where legacy TFVC projects/repos need to be reorganized across ADO collections.
    Process:
    1. Converts the TFVC source to Git (via git-tfs) or clones an existing Git repo
    2. Creates the target ADO project when moving an entire TFVC project by name
    3. Creates a new Git repo in the target ADO collection/project
    4. Pushes the converted repo to the target

    Supports an -Interactive mode that walks you through selecting the source project,
    repo, or folder and the target collection/project via numbered menus.

.PARAMETER ConfigPath
    Path to migration-config.json.

.PARAMETER Interactive
    Launch interactive mode — browse and select source and target via menus.

.PARAMETER SourceCollection
    Source ADO collection name (skipped in interactive mode).

.PARAMETER SourceServerUrl
    Optional source ADO server URL override. If omitted, uses sourceAdoServerUrl
    from config, then falls back to legacy adoServerUrl.

.PARAMETER SourceProject
    Source team project name (skipped in interactive mode).

.PARAMETER SourceRepoType
    Source repository type. TFVC uses git-tfs conversion; Git performs a mirror clone.

.PARAMETER TfvcPath
    TFVC path to migrate (e.g., $/SourceProject/AppFolder).

.PARAMETER SourceRepoName
    Source Git repository name when SourceRepoType is Git.

.PARAMETER MoveProjectByName
    Move all Git repositories in SourceProject to the target collection/project.
    This is a project-level move mode and does not require TFVC conversion.

.PARAMETER TargetCollection
    Target ADO collection name.

.PARAMETER TargetServerUrl
    Optional target ADO server URL override. If omitted, uses targetAdoServerUrl
    from config, then falls back to legacy adoServerUrl.

.PARAMETER TargetProject
    Target team project name. When MoveProjectByName is used and TargetProject is omitted,
    the script uses SourceProject and creates it in the target collection if needed.

.PARAMETER TargetRepoName
    Name for the new Git repo in the target project. When MoveProjectByName is used and
    TargetRepoName is omitted, the script uses SourceProject.

.PARAMETER HistoryDepth
    Number of changesets to include. Null = full history.

.PARAMETER SkipTargetRepoCreation
    Skip creating the target repo (if it already exists).

.EXAMPLE
    # Interactive mode — guided menus
    ./Move-RepoToCollection.ps1 -ConfigPath ./config/migration-config.json -Interactive

.EXAMPLE
    # Direct mode — all parameters specified
    ./Move-RepoToCollection.ps1 -ConfigPath ./config/migration-config.json `
        -SourceCollection "GAMS" -SourceProject "LegacyApp" -TfvcPath "$/LegacyApp" `
        -TargetCollection "ModernApps" -TargetProject "Platform" -TargetRepoName "legacy-app"

.EXAMPLE
    # Move an existing Git repo between collections without TFVC conversion
    ./Move-RepoToCollection.ps1 -ConfigPath ./config/migration-config.json `
        -SourceCollection "Legacy" -SourceProject "Apps" -SourceRepoType Git -SourceRepoName "billing-api" `
        -TargetCollection "GAMS-GIT-Repos" -TargetProject "Apps" -TargetRepoName "billing-api"

.EXAMPLE
    # Move an entire TFVC project by name
    ./Move-RepoToCollection.ps1 -ConfigPath ./config/migration-config.json `
    -SourceCollection "GAMS" -SourceProject "LegacyApp" -MoveProjectByName `
    -TargetCollection "ModernApps"
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory)]
    [string]$ConfigPath,

    [switch]$Interactive,

    [string]$SourceCollection,
    [string]$SourceProject,
    [ValidateSet('TFVC', 'Git')]
    [string]$SourceRepoType = 'TFVC',
    [string]$TfvcPath,
    [string]$SourceRepoName,
    [switch]$MoveProjectByName,
    [string]$SourceServerUrl,
    [string]$TargetCollection,
    [string]$TargetProject,
    [string]$TargetRepoName,
    [string]$TargetServerUrl,

    [int]$HistoryDepth,

    [switch]$SkipConversion,

    [switch]$NonInteractive,

    [int]$TimeoutMinutes,

    [int]$StallTimeoutMinutes,

    [switch]$SkipTargetRepoCreation
)

$ErrorActionPreference = 'Stop'
Import-Module "$PSScriptRoot/modules/AdoTfvcMigrator.psm1" -Force

function Get-GitBasicAuthHeaderValue {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$Pat
    )

    # Azure DevOps Git endpoints are more reliable with a non-empty username.
    $token = [Convert]::ToBase64String([Text.Encoding]::ASCII.GetBytes("ado:$Pat"))
    return "Authorization: Basic $token"
}

function Normalize-GitRemoteUrl {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$Url
    )

    # Some on-prem ADO URLs include literal spaces in project names; Git expects encoded URLs.
    return ($Url -replace ' ', '%20')
}

function Get-AdoProjectDetail {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$ServerUrl,
        [Parameter(Mandatory)]
        [string]$Collection,
        [Parameter(Mandatory)]
        [string]$ProjectName,
        [Parameter(Mandatory)]
        [string]$Pat,
        [switch]$IncludeCapabilities
    )

    $encodedProject = [Uri]::EscapeDataString($ProjectName)
    $url = "$ServerUrl/$Collection/_apis/projects/$encodedProject"
    if ($IncludeCapabilities) {
        $url += '?includeCapabilities=true'
    }

    return Invoke-AdoApi -Url $url -Pat $Pat -ApiVersion '7.1'
}

function Wait-AdoOperation {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$OperationUrl,
        [Parameter(Mandatory)]
        [string]$Pat,
        [int]$TimeoutMinutes = 20
    )

    $deadline = (Get-Date).AddMinutes($TimeoutMinutes)
    do {
        $operation = Invoke-AdoApi -Url $OperationUrl -Pat $Pat -ApiVersion '7.1'
        switch ($operation.status) {
            'succeeded' { return $operation }
            'failed' { throw "ADO operation failed: $($operation.resultMessage ?? $operation.detailedMessage ?? 'Unknown error')" }
            'cancelled' { throw "ADO operation was cancelled." }
        }

        Start-Sleep -Seconds 5
    }
    while ((Get-Date) -lt $deadline)

    throw "Timed out waiting for ADO operation to finish: $OperationUrl"
}

function Ensure-AdoProjectExists {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$ServerUrl,
        [Parameter(Mandatory)]
        [string]$Collection,
        [Parameter(Mandatory)]
        [string]$ProjectName,
        [Parameter(Mandatory)]
        [string]$Pat,
        [Parameter(Mandatory)]
        [object]$SourceProjectDetails,
        [string]$LogFile,
        [int]$TimeoutMinutes = 20
    )

    $projects = @(Get-AdoProjects -ServerUrl $ServerUrl -Collection $Collection -Pat $Pat)
    $existingProject = $projects | Where-Object { $_.name -eq $ProjectName } | Select-Object -First 1
    if ($existingProject) {
        return $existingProject
    }

    $templateTypeId = $SourceProjectDetails.capabilities.processTemplate.templateTypeId
    if (-not $templateTypeId) {
        $templateTypeId = '6b724908-ef14-45cf-84f8-768b5384da45'
    }

    $createBody = @{
        name = $ProjectName
        description = ($SourceProjectDetails.description ?? "Migrated from $($SourceProjectDetails.name)")
        capabilities = @{
            versioncontrol = @{ sourceControlType = 'Git' }
            processTemplate = @{ templateTypeId = $templateTypeId }
        }
    }

    Write-MigrationLog -Message "Creating target project '$ProjectName' in collection '$Collection'" -LogFile $LogFile -Level INFO
    $operation = Invoke-AdoApi -Url "$ServerUrl/$Collection/_apis/projects" -Pat $Pat -Method POST -Body $createBody -ApiVersion '7.1'
    Wait-AdoOperation -OperationUrl $operation.url -Pat $Pat -TimeoutMinutes $TimeoutMinutes | Out-Null
    Write-MigrationLog -Message "Created target project '$ProjectName'" -LogFile $LogFile -Level SUCCESS

    return Get-AdoProjectDetail -ServerUrl $ServerUrl -Collection $Collection -ProjectName $ProjectName -Pat $Pat
}

# ─── Load Config ───────────────────────────────────────────────────────────────

$config = Read-MigrationConfig -ConfigPath $ConfigPath
$logFile = Initialize-MigrationLog -LogDirectory $config.logDirectory -ScriptName 'MoveRepoToCollection'
$effectiveTimeoutMinutes = if ($TimeoutMinutes -gt 0) {
    $TimeoutMinutes
}
elseif ($config.migrationDefaults.timeoutMinutes -gt 0) {
    [int]$config.migrationDefaults.timeoutMinutes
}
else {
    120
}
if (-not $SourceServerUrl) {
    $SourceServerUrl = Get-ConfigAdoServerUrl -Config $config -Role Source
}
if (-not $TargetServerUrl) {
    $TargetServerUrl = Get-ConfigAdoServerUrl -Config $config -Role Target
}

# ─── Interactive Mode ──────────────────────────────────────────────────────────

if ($Interactive) {
    Show-MenuHeader -Title "Move TFVC Project or Repo to Another Collection"
    Write-Host "This wizard will walk you through selecting a source TFVC project, folder, or Git repo" -ForegroundColor DarkGray
    Write-Host "and a destination collection to move it to." -ForegroundColor DarkGray

    # ── 1. Pick source collection ──
    $SourceCollection = Select-AdoCollection -Config $config -Prompt 'Select SOURCE collection'
    if (-not $SourceCollection) { Write-Host "Cancelled." -ForegroundColor Yellow; return }

    # ── 2. Pick source project ──
    $SourceProject = Select-AdoProject -Config $config -Collection $SourceCollection -ServerUrl $SourceServerUrl -Prompt 'Select SOURCE project'
    if (-not $SourceProject) { Write-Host "Cancelled." -ForegroundColor Yellow; return }

    Show-MenuHeader -Title "Select SOURCE repo type"
    $repoTypeChoice = Show-NumberedMenu -Items @('Entire project (all Git repos)', 'TFVC folder', 'Git repository') -Prompt 'Select source type' -AllowBack
    if ($null -eq $repoTypeChoice) { Write-Host "Cancelled." -ForegroundColor Yellow; return }

    switch ($repoTypeChoice) {
        0 {
            $SourceRepoType = 'Git'
            $MoveProjectByName = $true
            $TfvcPath = $null
            $SourceRepoName = $null
        }
        1 {
            $SourceRepoType = 'TFVC'
            $currentPath = "`$/$SourceProject"
            while ($true) {
                $selected = Select-TfvcFolders -Config $config -Collection $SourceCollection `
                    -ProjectName $SourceProject -ServerUrl $SourceServerUrl -ParentPath $currentPath -Prompt 'Select a folder to move (or drill deeper)'
                if (-not $selected) { Write-Host "Cancelled." -ForegroundColor Yellow; return }

                $chosenPath = $selected[0]
                Write-Host ""
                Write-Host "  Selected: $chosenPath" -ForegroundColor Green
                Write-Host ""
                Write-Host "  [1] Use this folder as the source" -ForegroundColor White
                Write-Host "  [2] Drill into this folder to pick a subfolder" -ForegroundColor White
                Write-Host "  [0] Back" -ForegroundColor DarkGray
                Write-Host ""
                Write-Host "Choice: " -ForegroundColor Yellow -NoNewline
                $drillChoice = Read-Host

                switch ($drillChoice.Trim()) {
                    '1' { $TfvcPath = $chosenPath; break }
                    '2' { $currentPath = $chosenPath; continue }
                    '0' { Write-Host "Cancelled." -ForegroundColor Yellow; return }
                    default { $TfvcPath = $chosenPath; break }
                }
                if ($TfvcPath) { break }
            }
        }
        2 {
            $SourceRepoType = 'Git'
            $sourceRepo = Select-AdoGitRepo -Config $config -Collection $SourceCollection -ProjectName $SourceProject -ServerUrl $SourceServerUrl -Prompt 'Select SOURCE Git repo'
            if (-not $sourceRepo) { Write-Host "Cancelled." -ForegroundColor Yellow; return }
            $SourceRepoName = $sourceRepo.name
        }
    }

    # ── 4. Pick target collection ──
    Show-MenuHeader -Title "Select DESTINATION"
    if ($MoveProjectByName) {
        Write-Host "  Source: $SourceCollection / $SourceProject (all Git repos)" -ForegroundColor DarkGray
    }
    elseif ($SourceRepoType -eq 'Git') {
        Write-Host "  Source: $SourceCollection / $SourceProject / $SourceRepoName" -ForegroundColor DarkGray
    }
    else {
        Write-Host "  Source: $SourceCollection / $SourceProject / $TfvcPath" -ForegroundColor DarkGray
    }
    Write-Host ""

    $TargetCollection = Select-AdoCollection -Config $config -Prompt 'Select TARGET collection'
    if (-not $TargetCollection) { Write-Host "Cancelled." -ForegroundColor Yellow; return }

    # ── 5. Pick target project ──
    if ($MoveProjectByName) {
        $TargetProject = $SourceProject
        Write-Host "  Target project will be created or reused as: $TargetProject" -ForegroundColor DarkGray
    }
    else {
        $TargetProject = Select-AdoProject -Config $config -Collection $TargetCollection -ServerUrl $TargetServerUrl -Prompt 'Select TARGET project'
        if (-not $TargetProject) { Write-Host "Cancelled." -ForegroundColor Yellow; return }
    }

    # ── 6. Repo name ──
    if ($SourceRepoType -eq 'Git') {
        $defaultName = $SourceRepoName
    }
    else {
        $defaultName = ($TfvcPath -replace '^\$/', '' -replace '/', '-').ToLower()
    }
    if (-not $MoveProjectByName) {
        Write-Host ""
        Write-Host "  Git repo name [$defaultName]: " -ForegroundColor Yellow -NoNewline
        $nameInput = Read-Host
        $TargetRepoName = if ($nameInput.Trim()) { $nameInput.Trim() } else { $defaultName }
    }

    # ── 7. History depth ──
    if ($SourceRepoType -eq 'TFVC' -and -not $MoveProjectByName) {
        Write-Host "  History depth (enter for full history, or a number): " -ForegroundColor Yellow -NoNewline
        $depthInput = Read-Host
        if ($depthInput.Trim() -match '^\d+$') {
            $HistoryDepth = [int]$depthInput.Trim()
        }
    }

    # ── Confirm ──
    Show-MenuHeader -Title "Confirm Move"
    Write-Host "  Source Collection: $SourceCollection" -ForegroundColor White
    Write-Host "  Source Project:    $SourceProject" -ForegroundColor White
    if ($MoveProjectByName) {
        Write-Host "  Source Type:       Project (all Git repos)" -ForegroundColor White
        Write-Host "  Source Scope:      Entire project ($SourceProject)" -ForegroundColor White
    }
    else {
        Write-Host "  Source Type:       $SourceRepoType" -ForegroundColor White
    }

    if (-not $MoveProjectByName -and $SourceRepoType -eq 'Git') {
        Write-Host "  Source Repo:       $SourceRepoName" -ForegroundColor White
    }
    elseif (-not $MoveProjectByName) {
        Write-Host "  TFVC Path:         $TfvcPath" -ForegroundColor White
    }
    Write-Host "  Target Collection: $TargetCollection" -ForegroundColor White
    Write-Host "  Target Project:    $TargetProject" -ForegroundColor White
    if ($MoveProjectByName) {
        Write-Host "  Target Git Repos:  All repos from source project (same names)" -ForegroundColor White
    }
    else {
        Write-Host "  Target Git Repo:   $TargetRepoName" -ForegroundColor White
    }
    if ($SourceRepoType -eq 'TFVC' -and -not $MoveProjectByName) {
        if ($HistoryDepth) {
            Write-Host "  History Depth:     $HistoryDepth changesets" -ForegroundColor White
        }
        else {
            Write-Host "  History Depth:     Full" -ForegroundColor White
        }
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

if ($MoveProjectByName) {
    # Project move mode always uses Git repository mirroring.
    $SourceRepoType = 'Git'
    if (-not $TargetProject) {
        $TargetProject = $SourceProject
    }
}

if (-not $SourceCollection) { throw "SourceCollection is required. Use -Interactive or provide -SourceCollection." }
if (-not $SourceProject)    { throw "SourceProject is required. Use -Interactive or provide -SourceProject." }
if (-not $TargetCollection) { throw "TargetCollection is required. Use -Interactive or provide -TargetCollection." }
if (-not $TargetProject)    { throw "TargetProject is required. Use -Interactive or provide -TargetProject." }
if ((-not $MoveProjectByName) -and (-not $TargetRepoName)) {
    throw "TargetRepoName is required. Use -Interactive or provide -TargetRepoName."
}

switch ($SourceRepoType) {
    'TFVC' {
        if (-not $TfvcPath) {
            throw "TfvcPath is required when SourceRepoType is TFVC. Use -Interactive or provide -TfvcPath."
        }
    }
    'Git' {
        if ((-not $MoveProjectByName) -and (-not $Interactive) -and (-not $SourceRepoName)) {
            throw "SourceRepoName is required when SourceRepoType is Git. Use -Interactive or provide -SourceRepoName."
        }
    }
}

$sourceDescriptor = if ($MoveProjectByName) {
    "$SourceCollection/$SourceProject (all Git repos)"
}
elseif ($SourceRepoType -eq 'Git') {
    "$SourceCollection/$SourceProject/$SourceRepoName"
}
else {
    "$SourceCollection/$SourceProject ($TfvcPath)"
}

Write-MigrationLog -Message "Starting cross-collection move" -LogFile $logFile -Level INFO
Write-MigrationLog -Message "  Source URL: $SourceServerUrl" -LogFile $logFile
Write-MigrationLog -Message "  Source Type: $SourceRepoType" -LogFile $logFile
Write-MigrationLog -Message "  Source: $sourceDescriptor" -LogFile $logFile
Write-MigrationLog -Message "  Target URL: $TargetServerUrl" -LogFile $logFile
if ($MoveProjectByName) {
    Write-MigrationLog -Message "  Target: $TargetCollection/$TargetProject (all repos from source project)" -LogFile $logFile
}
else {
    Write-MigrationLog -Message "  Target: $TargetCollection/$TargetProject/$TargetRepoName" -LogFile $logFile
}
Write-MigrationLog -Message "  Timeout: $effectiveTimeoutMinutes minute(s)" -LogFile $logFile

# ─── Validate Source & Target ──────────────────────────────────────────────────

$sourceConfig = $config.collections[$SourceCollection]
$targetConfig = $config.collections[$TargetCollection]

if (-not $sourceConfig) { throw "Source collection '$SourceCollection' not in config." }
if (-not $targetConfig) { throw "Target collection '$TargetCollection' not in config." }

$sourcePat = $sourceConfig.pat
$targetPat = $targetConfig.pat

# Test source
$sourceTest = Test-AdoConnection -ServerUrl $SourceServerUrl -Collection $SourceCollection -Pat $sourcePat
if (-not $sourceTest.Connected) {
    throw "Cannot connect to source collection '$SourceCollection': $($sourceTest.Error)"
}

# Test target
$targetTest = Test-AdoConnection -ServerUrl $TargetServerUrl -Collection $TargetCollection -Pat $targetPat
if (-not $targetTest.Connected) {
    throw "Cannot connect to target collection '$TargetCollection': $($targetTest.Error)"
}

if ($SourceProject -notin $sourceTest.Projects) {
    throw "Project '$SourceProject' not found in collection '$SourceCollection'. Available: $($sourceTest.Projects -join ', ')"
}

if ($MoveProjectByName) {
    $sourceProjectDetails = Get-AdoProjectDetail -ServerUrl $SourceServerUrl -Collection $SourceCollection -ProjectName $SourceProject -Pat $sourcePat -IncludeCapabilities
    if ($TargetProject -notin $targetTest.Projects) {
        Ensure-AdoProjectExists -ServerUrl $TargetServerUrl -Collection $TargetCollection -ProjectName $TargetProject -Pat $targetPat -SourceProjectDetails $sourceProjectDetails -LogFile $logFile -TimeoutMinutes $effectiveTimeoutMinutes | Out-Null
        $targetTest = Test-AdoConnection -ServerUrl $TargetServerUrl -Collection $TargetCollection -Pat $targetPat
    }
}
elseif ($TargetProject -notin $targetTest.Projects) {
    throw "Project '$TargetProject' not found in collection '$TargetCollection'. Available: $($targetTest.Projects -join ', ')"
}

$projectRepos = @()
if ($MoveProjectByName) {
    $projectRepos = @(Get-AdoGitRepositories -ServerUrl $SourceServerUrl -Collection $SourceCollection -ProjectName $SourceProject -Pat $sourcePat)
    if ($projectRepos.Count -eq 0) {
        throw "No Git repositories were found in source project '$SourceCollection/$SourceProject'. Project move mode migrates Git repos only."
    }
    Write-MigrationLog -Message "Found $($projectRepos.Count) Git repo(s) in source project to move" -LogFile $logFile -Level INFO
}

if ($SourceRepoType -eq 'Git') {
    $sourceRepos = @(Get-AdoGitRepositories -ServerUrl $SourceServerUrl -Collection $SourceCollection -ProjectName $SourceProject -Pat $sourcePat)
    $sourceRepo = $sourceRepos | Where-Object { $_.name -eq $SourceRepoName } | Select-Object -First 1
    if (-not $sourceRepo) {
        throw "Git repo '$SourceRepoName' not found in '$SourceCollection/$SourceProject'. Available: $($sourceRepos.name -join ', ')"
    }
}

Write-MigrationLog -Message "Source and target connections verified" -LogFile $logFile -Level SUCCESS

# ─── Project move mode (all Git repos) ────────────────────────────────────────

if ($MoveProjectByName) {
    Write-MigrationLog -Message "Step 1-3: Moving all Git repos from '$SourceProject' to '$TargetCollection/$TargetProject'" -LogFile $logFile -Level INFO

    $movedCount = 0
    foreach ($repo in $projectRepos) {
        $repoName = $repo.name
        $localRepoPath = Join-Path $config.outputDirectory $repoName

        Write-MigrationLog -Message "Moving repo '$repoName'" -LogFile $logFile -Level INFO

        if (Test-Path $localRepoPath) {
            Write-MigrationLog -Message "Removing existing local repo path: $localRepoPath" -LogFile $logFile -Level WARN
            Remove-Item -Recurse -Force $localRepoPath
        }

        $sourceRemoteUrl = Normalize-GitRemoteUrl -Url $repo.remoteUrl
        $sourceAuthHeader = Get-GitBasicAuthHeaderValue -Pat $sourcePat
        Invoke-Git -Arguments "-c http.extraHeader=`"$sourceAuthHeader`" clone --mirror `"$sourceRemoteUrl`" `"$localRepoPath`"" -LogFile $logFile

        if (-not (Test-Path (Join-Path $localRepoPath 'HEAD'))) {
            throw "Clone failed for '$repoName' — no mirrored Git repo found at $localRepoPath"
        }

        $projects = Get-AdoProjects -ServerUrl $TargetServerUrl -Collection $TargetCollection -Pat $targetPat
        $targetProjectObj = $projects | Where-Object { $_.name -eq $TargetProject }
        if (-not $targetProjectObj) {
            throw "Project '$TargetProject' not found in '$TargetCollection'."
        }

        $createBody = @{
            name    = $repoName
            project = @{ id = $targetProjectObj.id }
        }

        try {
            $url = "$TargetServerUrl/$TargetCollection/_apis/git/repositories"
            $newRepo = Invoke-AdoApi -Url $url -Pat $targetPat -Method POST -Body $createBody
            $remoteUrl = $newRepo.remoteUrl
            Write-MigrationLog -Message "Created target repo '$repoName'" -LogFile $logFile -Level SUCCESS
        }
        catch {
            if ($_.Exception.Message -like '*already exists*' -or $_.Exception.Message -like '*409*') {
                Write-MigrationLog -Message "Target repo '$repoName' already exists — will push to existing" -LogFile $logFile -Level WARN
                $url = "$TargetServerUrl/$TargetCollection/$TargetProject/_apis/git/repositories/$repoName"
                $existingRepo = Invoke-AdoApi -Url $url -Pat $targetPat
                $remoteUrl = $existingRepo.remoteUrl
            }
            else {
                throw
            }
        }

        $normalizedTargetRemoteUrl = Normalize-GitRemoteUrl -Url $remoteUrl
        $targetAuthHeader = Get-GitBasicAuthHeaderValue -Pat $targetPat

        Invoke-Git -Arguments "remote add target `"$normalizedTargetRemoteUrl`"" -WorkingDirectory $localRepoPath -LogFile $logFile
        Invoke-Git -Arguments "-c http.extraHeader=`"$targetAuthHeader`" push --mirror target" -WorkingDirectory $localRepoPath -LogFile $logFile
        Invoke-Git -Arguments "remote remove target" -WorkingDirectory $localRepoPath -LogFile $logFile

        $movedCount++
        Write-MigrationLog -Message "Moved '$repoName' ($movedCount/$($projectRepos.Count))" -LogFile $logFile -Level SUCCESS
    }

    Write-Host ""
    Write-MigrationLog -Message "Cross-collection project move complete!" -LogFile $logFile -Level SUCCESS
    Write-MigrationLog -Message "  Source:     $SourceCollection/$SourceProject" -LogFile $logFile -Level SUCCESS
    Write-MigrationLog -Message "  Target:     $TargetCollection/$TargetProject" -LogFile $logFile -Level SUCCESS
    Write-MigrationLog -Message "  Repos moved: $movedCount" -LogFile $logFile -Level SUCCESS
    Write-MigrationLog -Message "  Log:        $logFile" -LogFile $logFile -Level INFO
    return
}

# ─── Step 1: Convert TFVC to Git ──────────────────────────────────────────────

$localRepoPath = Join-Path $config.outputDirectory $TargetRepoName

if ($SkipConversion) {
    Write-MigrationLog -Message "Step 1: Skipping conversion (already done upstream)" -LogFile $logFile -Level INFO
}
elseif ($SourceRepoType -eq 'Git') {
    Write-MigrationLog -Message "Step 1: Cloning source Git repo '$SourceRepoName'" -LogFile $logFile -Level INFO

    if (Test-Path $localRepoPath) {
        Write-MigrationLog -Message "Removing existing local repo path: $localRepoPath" -LogFile $logFile -Level WARN
        Remove-Item -Recurse -Force $localRepoPath
    }

    $sourceRemoteUrl = Normalize-GitRemoteUrl -Url $sourceRepo.remoteUrl
    $sourceAuthHeader = Get-GitBasicAuthHeaderValue -Pat $sourcePat
    Invoke-Git -Arguments "-c http.extraHeader=`"$sourceAuthHeader`" clone --mirror `"$sourceRemoteUrl`" `"$localRepoPath`"" -LogFile $logFile

    Write-MigrationLog -Message "Git clone complete" -LogFile $logFile -Level SUCCESS
}
else {
    Write-MigrationLog -Message "Step 1: Converting TFVC to Git" -LogFile $logFile -Level INFO

    $convertParams = @{
        ConfigPath     = $ConfigPath
        ServerUrl      = $SourceServerUrl
        Collection     = $SourceCollection
        ProjectName    = $SourceProject
        TfvcPath       = $TfvcPath
        OutputRepoName = $TargetRepoName
    }
    if ($HistoryDepth) {
        $convertParams.HistoryDepth = $HistoryDepth
    }
    if ($NonInteractive) {
        $convertParams.NonInteractive = $true
    }
    if ($TimeoutMinutes) {
        $convertParams.TimeoutMinutes = $TimeoutMinutes
    }
    if ($StallTimeoutMinutes) {
        $convertParams.StallTimeoutMinutes = $StallTimeoutMinutes
    }

    & "$PSScriptRoot/Convert-TfvcToGit.ps1" @convertParams

    Write-MigrationLog -Message "Conversion complete" -LogFile $logFile -Level SUCCESS
}

if ($SourceRepoType -eq 'Git') {
    if (-not (Test-Path (Join-Path $localRepoPath 'HEAD'))) {
        throw "Clone failed — no mirrored Git repo found at $localRepoPath"
    }
}
elseif (-not $SkipConversion -and -not (Test-Path (Join-Path $localRepoPath '.git'))) {
    throw "Conversion failed — no Git repo found at $localRepoPath"
}

# ─── Step 2: Create target Git repo in ADO ─────────────────────────────────────

if (-not $SkipTargetRepoCreation) {
    Write-MigrationLog -Message "Step 2: Creating Git repo '$TargetRepoName' in $TargetCollection/$TargetProject" -LogFile $logFile -Level INFO

    # Find the project ID
    $projects = Get-AdoProjects -ServerUrl $TargetServerUrl -Collection $TargetCollection -Pat $targetPat
    $targetProjectObj = $projects | Where-Object { $_.name -eq $TargetProject }
    if (-not $targetProjectObj) {
        throw "Project '$TargetProject' not found in '$TargetCollection'."
    }

    $createBody = @{
        name    = $TargetRepoName
        project = @{
            id = $targetProjectObj.id
        }
    }

    try {
        $url = "$TargetServerUrl/$TargetCollection/_apis/git/repositories"
        $newRepo = Invoke-AdoApi -Url $url -Pat $targetPat -Method POST -Body $createBody
        $remoteUrl = $newRepo.remoteUrl
        Write-MigrationLog -Message "Created repo: $remoteUrl" -LogFile $logFile -Level SUCCESS
    }
    catch {
        if ($_.Exception.Message -like '*already exists*' -or $_.Exception.Message -like '*409*') {
            Write-MigrationLog -Message "Repo already exists — will push to existing" -LogFile $logFile -Level WARN
            # Fetch existing repo URL
            $url = "$TargetServerUrl/$TargetCollection/$TargetProject/_apis/git/repositories/$TargetRepoName"
            $existingRepo = Invoke-AdoApi -Url $url -Pat $targetPat
            $remoteUrl = $existingRepo.remoteUrl
        }
        else {
            throw
        }
    }
}
else {
    # Get existing repo URL
    $url = "$TargetServerUrl/$TargetCollection/$TargetProject/_apis/git/repositories/$TargetRepoName"
    $existingRepo = Invoke-AdoApi -Url $url -Pat $targetPat
    $remoteUrl = $existingRepo.remoteUrl
}

# ─── Step 3: Push to target ────────────────────────────────────────────────────

Write-MigrationLog -Message "Step 3: Pushing to target repo" -LogFile $logFile -Level INFO

# Normalize URL and use auth header (avoid PAT-in-URL parsing edge cases).
$normalizedTargetRemoteUrl = Normalize-GitRemoteUrl -Url $remoteUrl
$targetAuthHeader = Get-GitBasicAuthHeaderValue -Pat $targetPat

# Add remote and push
Invoke-Git -Arguments "remote add target `"$normalizedTargetRemoteUrl`"" -WorkingDirectory $localRepoPath -LogFile $logFile

$defaultBranch = $config.defaultBranch ?? 'main'
if ($SourceRepoType -eq 'Git') {
    Invoke-Git -Arguments "-c http.extraHeader=`"$targetAuthHeader`" push --mirror target" -WorkingDirectory $localRepoPath -LogFile $logFile
}
else {
    Invoke-Git -Arguments "-c http.extraHeader=`"$targetAuthHeader`" push target $defaultBranch --force" -WorkingDirectory $localRepoPath -LogFile $logFile
    Invoke-Git -Arguments "-c http.extraHeader=`"$targetAuthHeader`" push target --tags" -WorkingDirectory $localRepoPath -LogFile $logFile
}

# Remove remote with PAT from local config
Invoke-Git -Arguments "remote remove target" -WorkingDirectory $localRepoPath -LogFile $logFile

Write-MigrationLog -Message "Push complete!" -LogFile $logFile -Level SUCCESS

# ─── Summary ──────────────────────────────────────────────────────────────────

Write-Host ""
Write-MigrationLog -Message "Cross-collection move complete!" -LogFile $logFile -Level SUCCESS
Write-MigrationLog -Message "  Source:     $sourceDescriptor" -LogFile $logFile -Level SUCCESS
Write-MigrationLog -Message "  Target:     $TargetCollection/$TargetProject/$TargetRepoName" -LogFile $logFile -Level SUCCESS
Write-MigrationLog -Message "  Remote URL: $remoteUrl" -LogFile $logFile -Level SUCCESS
Write-MigrationLog -Message "  Local copy: $localRepoPath" -LogFile $logFile -Level INFO
Write-MigrationLog -Message "  Log:        $logFile" -LogFile $logFile -Level INFO
