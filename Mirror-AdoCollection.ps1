#Requires -Version 7.0
<#
.SYNOPSIS
    Mirrors every project and Git repository from one ADO collection to another,
    preserving names. Designed for on-prem ADO Server -> Azure DevOps Services
    organisation migrations (e.g. GAMS-GIT-Repos -> MDR-GAMS-ADO), but works for
    any source/target combination defined in the config.

.DESCRIPTION
    For each project in the SOURCE collection:
      1. Ensure the project exists in the TARGET (auto-create if missing).
      2. For each Git repo in the source project, mirror it to the target
         project of the same name using `git clone --mirror` followed by
         `git push --mirror`. All branches, tags, and notes are preserved.

    Cross-server is supported via an optional per-collection `serverUrl` entry
    in the config. For an Azure DevOps Services org, set:
        "serverUrl": "https://dev.azure.com"
    and use the org name as the collection key (e.g. "MDR-GAMS-ADO").

    The script writes a CSV manifest (and a final JSON report) under
    outputDirectory. Re-running with -ResumeManifest skips rows that have
    status Success so partial runs can be safely continued.

.PARAMETER ConfigPath
    Path to migration-config.json.

.PARAMETER SourceCollection
    Source collection key as defined under config.collections.
    Required unless -Interactive.

.PARAMETER TargetCollection
    Target collection key as defined under config.collections. For an Azure
    DevOps Services org, this is the org name. Required unless -Interactive.

.PARAMETER Interactive
    Walk through source/target/filter selection via numbered menus.

.PARAMETER IncludeProjects
    Optional list of project names to include. If empty, ALL source projects
    are processed.

.PARAMETER ExcludeProjects
    Optional list of project names to skip.

.PARAMETER SourceProject
    Convenience single-project filter. Equivalent to passing
    `-IncludeProjects <name>`. Intended for callers that already picked a
    single project to migrate (see Invoke-ProjectMigration.ps1).

.PARAMETER WorkingDirectory
    Where to put temporary bare clones. Defaults to <outputDirectory>/mirror-cache.

.PARAMETER DryRun
    Show what would happen without creating anything in the target or pushing.

.PARAMETER PreviewOnly
    Write the preview CSV and stop without doing any work.

.PARAMETER ResumeManifest
    Path to an existing manifest CSV from a prior run. Repos with Status=Success
    are skipped; everything else is re-attempted.

.PARAMETER Force
    Suppress the interactive "type 'yes' to continue" confirmation prompt.

.PARAMETER KeepCache
    Keep the local bare clone after a successful push (default: delete on success).

.PARAMETER ProcessTemplateName
    Process template to use when auto-creating target projects (default: Agile).

.PARAMETER TimeoutMinutes
    Hard timeout per git clone/push (0 = no timeout). Defaults to config.migrationDefaults.timeoutMinutes.

.PARAMETER StallTimeoutMinutes
    Abort a git operation if no output appears for this many minutes
    (0 = disabled). Defaults to config.migrationDefaults.stallTimeoutMinutes.

.EXAMPLE
    # Interactive — pick source/target, preview, then confirm
    ./Mirror-AdoCollection.ps1 -ConfigPath ./config/migration-config.json -Interactive

.EXAMPLE
    # On-prem ADO 2022 collection -> Azure DevOps Services org, full mirror
    ./Mirror-AdoCollection.ps1 `
        -ConfigPath ./config/migration-config.json `
        -SourceCollection GAMS-GIT-Repos `
        -TargetCollection MDR-GAMS-ADO `
        -Force

.EXAMPLE
    # Preview-only: just write the manifest CSV, do nothing else
    ./Mirror-AdoCollection.ps1 -ConfigPath ./config/migration-config.json `
        -SourceCollection GAMS-GIT-Repos -TargetCollection MDR-GAMS-ADO -PreviewOnly

.EXAMPLE
    # Resume from a partial run
    ./Mirror-AdoCollection.ps1 -ConfigPath ./config/migration-config.json `
        -SourceCollection GAMS-GIT-Repos -TargetCollection MDR-GAMS-ADO `
        -ResumeManifest ./output/mirror-MANIFEST-20260520-093011.csv -Force
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory)]
    [string]$ConfigPath,

    [string]$SourceCollection,
    [string]$TargetCollection,

    [switch]$Interactive,

    [string[]]$IncludeProjects,
    [string[]]$ExcludeProjects,

    # Convenience: single-project filter. Merged into IncludeProjects.
    [string]$SourceProject,

    [string]$WorkingDirectory,

    [switch]$DryRun,
    [switch]$PreviewOnly,

    [string]$ResumeManifest,

    [switch]$Force,
    [switch]$KeepCache,

    [string]$ProcessTemplateName = 'Agile',

    [int]$TimeoutMinutes = -1,
    [int]$StallTimeoutMinutes = -1,

    # Suppress nested Write-Progress bars (handy for CI / non-interactive hosts).
    # Per-repo scrollback ticks are still printed.
    [switch]$NoProgress
)

$ErrorActionPreference = 'Stop'
Import-Module "$PSScriptRoot/modules/AdoTfvcMigrator.psm1" -Force

# ─── Load Config ───────────────────────────────────────────────────────────────

$config = Read-MigrationConfig -ConfigPath $ConfigPath
$logFile = Initialize-MigrationLog -LogDirectory $config.logDirectory -ScriptName 'MirrorAdoCollection'
$timestamp = Get-Date -Format 'yyyyMMdd-HHmmss'

# Default timeouts from config if not explicitly set
if ($TimeoutMinutes -lt 0) {
    $TimeoutMinutes = [int]($config.migrationDefaults.timeoutMinutes ?? 0)
}
if ($StallTimeoutMinutes -lt 0) {
    $StallTimeoutMinutes = [int]($config.migrationDefaults.stallTimeoutMinutes ?? 0)
}

# ─── Interactive Selection ────────────────────────────────────────────────────

if ($Interactive) {
    Show-MenuHeader -Title 'Mirror Git Collection (Source ADO -> Target ADO)'
    Write-Host "Mirrors every project and Git repository from a source collection to a" -ForegroundColor DarkGray
    Write-Host "target collection. Cross-server (on-prem -> Azure DevOps Services) supported." -ForegroundColor DarkGray
    Write-Host ""

    if (-not $SourceCollection) {
        $SourceCollection = Select-AdoCollection -Config $config -Prompt 'Select SOURCE collection'
        if (-not $SourceCollection) { Write-Host 'Cancelled.' -ForegroundColor Yellow; return }
    }

    if (-not $TargetCollection) {
        $TargetCollection = Select-AdoCollection -Config $config -Prompt 'Select TARGET collection'
        if (-not $TargetCollection) { Write-Host 'Cancelled.' -ForegroundColor Yellow; return }
    }
}

if (-not $SourceCollection -or -not $TargetCollection) {
    throw 'Both -SourceCollection and -TargetCollection are required (or use -Interactive).'
}
if ($SourceCollection -eq $TargetCollection) {
    throw "Source and target collections must differ ('$SourceCollection')."
}

# ─── Resolve per-collection endpoints + PATs ──────────────────────────────────

$sourceUrl = Get-CollectionServerUrl -Config $config -Collection $SourceCollection
$targetUrl = Get-CollectionServerUrl -Config $config -Collection $TargetCollection
$sourcePat = Get-CollectionPat        -Config $config -Collection $SourceCollection
$targetPat = Get-CollectionPat        -Config $config -Collection $TargetCollection

Write-MigrationLog -Message "Source: $sourceUrl/$SourceCollection" -LogFile $logFile
Write-MigrationLog -Message "Target: $targetUrl/$TargetCollection" -LogFile $logFile

# ─── Validate Both Endpoints Before Doing Any Work ────────────────────────────

Write-Host ""
Write-Host "Validating SOURCE connection..." -ForegroundColor Cyan
$srcCheck = Test-AdoConnection -ServerUrl $sourceUrl -Collection $SourceCollection -Pat $sourcePat
if (-not $srcCheck.Connected) {
    throw "Cannot reach source '$SourceCollection': $($srcCheck.Error)"
}
Write-Host "  ✓ Source OK ($($srcCheck.ProjectCount) projects)" -ForegroundColor Green

Write-Host "Validating TARGET connection..." -ForegroundColor Cyan
$tgtCheck = Test-AdoConnection -ServerUrl $targetUrl -Collection $TargetCollection -Pat $targetPat
if (-not $tgtCheck.Connected) {
    throw "Cannot reach target '$TargetCollection': $($tgtCheck.Error)"
}
Write-Host "  ✓ Target OK ($($tgtCheck.ProjectCount) existing projects)" -ForegroundColor Green

# ─── Enumerate Source Projects + Repos ────────────────────────────────────────

Write-Host ""
Write-Host "Enumerating source projects and Git repos..." -ForegroundColor Cyan
$srcProjects = Get-AdoProjects -ServerUrl $sourceUrl -Collection $SourceCollection -Pat $sourcePat | Sort-Object name

# Merge the convenience -SourceProject scalar into the IncludeProjects list
if ($SourceProject) {
    if ($IncludeProjects) {
        if ($IncludeProjects -notcontains $SourceProject) {
            $IncludeProjects = @($IncludeProjects) + $SourceProject
        }
    }
    else {
        $IncludeProjects = @($SourceProject)
    }
}

if ($IncludeProjects) {
    $srcProjects = $srcProjects | Where-Object { $_.name -in $IncludeProjects }
}
if ($ExcludeProjects) {
    $srcProjects = $srcProjects | Where-Object { $_.name -notin $ExcludeProjects }
}

if (-not $srcProjects) {
    Write-Host "No source projects matched the filter — nothing to do." -ForegroundColor Yellow
    return
}

# Interactive project filter
if ($Interactive -and -not $IncludeProjects) {
    Show-MenuHeader -Title "Source projects in '$SourceCollection' ($($srcProjects.Count) total)"
    Write-Host "  [1] Mirror ALL $($srcProjects.Count) projects" -ForegroundColor White
    Write-Host "  [2] Pick specific projects to include" -ForegroundColor White
    Write-Host "  [0] Cancel" -ForegroundColor DarkGray
    Write-Host ""
    Write-Host 'Choose: ' -ForegroundColor Yellow -NoNewline
    $filterChoice = Read-Host
    if ($filterChoice.Trim() -eq '0') { Write-Host 'Cancelled.' -ForegroundColor Yellow; return }
    if ($filterChoice.Trim() -eq '2') {
        $items = $srcProjects.name
        $indices = Show-NumberedMenu -Items $items -Prompt 'Select projects' -MultiSelect -AllowBack
        if ($null -eq $indices) { Write-Host 'Cancelled.' -ForegroundColor Yellow; return }
        $srcProjects = @($indices | ForEach-Object { $srcProjects[$_] })
    }
}

# Build the work list
$workItems = [System.Collections.Generic.List[object]]::new()
$projectsWithoutRepos = [System.Collections.Generic.List[string]]::new()

foreach ($srcProj in $srcProjects) {
    try {
        $repos = Get-AdoGitRepositories -ServerUrl $sourceUrl -Collection $SourceCollection `
            -ProjectIdOrName $srcProj.id -Pat $sourcePat
    }
    catch {
        Write-MigrationLog -Message "Could not list repos in source project '$($srcProj.name)': $($_.Exception.Message)" -LogFile $logFile -Level WARN
        $repos = @()
    }
    if (-not $repos -or $repos.Count -eq 0) {
        $projectsWithoutRepos.Add($srcProj.name)
        continue
    }
    foreach ($repo in $repos) {
        $workItems.Add([PSCustomObject]@{
                Timestamp        = ''
                SourceCollection = $SourceCollection
                SourceProject    = $srcProj.name
                SourceProjectId  = $srcProj.id
                SourceRepo       = $repo.name
                SourceRepoId     = $repo.id
                SourceRepoUrl    = $repo.remoteUrl
                SourceSize       = $repo.size
                TargetCollection = $TargetCollection
                TargetProject    = $srcProj.name
                TargetRepo       = $repo.name
                TargetRepoUrl    = ''
                Status           = 'Pending'
                Reason           = ''
                DurationSeconds  = 0
            })
    }
}

$workItems = $workItems.ToArray()

Write-Host ""
Write-Host "Discovered:" -ForegroundColor Cyan
Write-Host "  Projects   : $($srcProjects.Count)" -ForegroundColor White
Write-Host "  Repos      : $($workItems.Count)" -ForegroundColor White
if ($projectsWithoutRepos.Count -gt 0) {
    Write-Host "  Empty projects (no Git repos, will still create target project): $($projectsWithoutRepos.Count)" -ForegroundColor DarkGray
}

# ─── Apply Resume Manifest ────────────────────────────────────────────────────

$resumeStatus = @{}
if ($ResumeManifest) {
    if (-not (Test-Path $ResumeManifest)) {
        throw "ResumeManifest not found: $ResumeManifest"
    }
    $prev = Import-Csv $ResumeManifest
    foreach ($row in $prev) {
        $key = "$($row.SourceCollection)|$($row.SourceProject)|$($row.SourceRepo)"
        $resumeStatus[$key] = $row.Status
    }
    Write-Host ""
    Write-Host "Loaded resume manifest with $($prev.Count) prior rows. Repos with Status=Success will be skipped." -ForegroundColor Yellow
}

# ─── Write Preview Manifest ───────────────────────────────────────────────────

if (-not (Test-Path $config.outputDirectory)) {
    New-Item -ItemType Directory -Path $config.outputDirectory -Force | Out-Null
}

$previewPath = Join-Path $config.outputDirectory "mirror-PREVIEW-$timestamp.csv"
$workItems | Select-Object SourceCollection, SourceProject, SourceRepo, SourceRepoUrl, SourceSize,
TargetCollection, TargetProject, TargetRepo |
    Export-Csv -Path $previewPath -NoTypeInformation -Encoding utf8

Write-Host ""
Write-Host "Preview written: $previewPath" -ForegroundColor Cyan

if ($PreviewOnly) {
    Write-Host "Preview-only mode: stopping here. Open the CSV above, then re-run without -PreviewOnly." -ForegroundColor Yellow
    return
}

# ─── Confirm ──────────────────────────────────────────────────────────────────

$action = if ($DryRun) { 'DRY RUN' } else { 'MIRROR' }
Write-Host ""
Write-Host "About to $action $($workItems.Count) repo(s) across $($srcProjects.Count) project(s)" -ForegroundColor Yellow
Write-Host "  From: $sourceUrl/$SourceCollection" -ForegroundColor White
Write-Host "  To  : $targetUrl/$TargetCollection" -ForegroundColor White
if (-not $DryRun) {
    Write-Host ""
    Write-Host "Missing target projects will be CREATED automatically (template: $ProcessTemplateName)." -ForegroundColor Yellow
    Write-Host "Missing target repos will be CREATED automatically." -ForegroundColor Yellow
    Write-Host "Existing target repos will receive a force `git push --mirror` (refs WILL be overwritten)." -ForegroundColor Red
}

if (-not $Force) {
    Write-Host ""
    Write-Host "Type 'yes' to continue: " -ForegroundColor Yellow -NoNewline
    $confirm = Read-Host
    if ($confirm.Trim().ToLower() -ne 'yes') {
        Write-Host 'Cancelled.' -ForegroundColor Yellow
        return
    }
}

# ─── Prepare Working Directory ────────────────────────────────────────────────

if (-not $WorkingDirectory) {
    $WorkingDirectory = Join-Path $config.outputDirectory 'mirror-cache'
}
if (-not (Test-Path $WorkingDirectory)) {
    New-Item -ItemType Directory -Path $WorkingDirectory -Force | Out-Null
}

$manifestPath = Join-Path $config.outputDirectory "mirror-MANIFEST-$timestamp.csv"

# ─── Cache Target Project IDs (created on demand) ─────────────────────────────

$targetProjectCache = @{}

function Resolve-TargetProject {
    param(
        [Parameter(Mandatory)] [string]$ProjectName,
        [string]$Description
    )

    if ($targetProjectCache.ContainsKey($ProjectName)) {
        return $targetProjectCache[$ProjectName]
    }
    $existing = Get-AdoTeamProject -ServerUrl $targetUrl -Collection $TargetCollection `
        -Name $ProjectName -Pat $targetPat
    if ($existing) {
        $targetProjectCache[$ProjectName] = $existing
        return $existing
    }
    if ($DryRun) {
        Write-Host "  [DRY RUN] Would create target project '$ProjectName' (template: $ProcessTemplateName)" -ForegroundColor Yellow
        $stub = [PSCustomObject]@{ id = '00000000-0000-0000-0000-000000000000'; name = $ProjectName }
        $targetProjectCache[$ProjectName] = $stub
        return $stub
    }
    Write-Host "  Creating target project '$ProjectName' (template: $ProcessTemplateName)..." -ForegroundColor Cyan
    $created = New-AdoTeamProject -ServerUrl $targetUrl -Collection $TargetCollection `
        -Pat $targetPat -Name $ProjectName -Description ($Description ?? '') `
        -PreferredProcessName $ProcessTemplateName -LogFile $logFile
    $targetProjectCache[$ProjectName] = $created
    return $created
}

# ─── Progress Helpers ─────────────────────────────────────────────────────────

# IDs:  10 = overall (collection),  11 = current project,  12 = current repo step
$ProgressIdOverall = 10
$ProgressIdProject = 11
$ProgressIdStep    = 12

function Format-Duration {
    param([double]$Seconds)
    if ($Seconds -le 0 -or [double]::IsInfinity($Seconds) -or [double]::IsNaN($Seconds)) { return '--:--:--' }
    return [TimeSpan]::FromSeconds([int]$Seconds).ToString('hh\:mm\:ss')
}

function Write-OverallProgress {
    param(
        [int]$Completed,
        [int]$Total,
        [int]$SuccessCount,
        [int]$SkipCount,
        [int]$FailCount,
        [string]$CurrentProject,
        [int]$ProjIndex,
        [int]$ProjTotal,
        [System.Diagnostics.Stopwatch]$OverallStopwatch
    )
    if ($script:NoProgress) { return }
    $pct = if ($Total -gt 0) { [int](($Completed * 100) / $Total) } else { 0 }
    $elapsedSec = $OverallStopwatch.Elapsed.TotalSeconds
    $avg = if ($Completed -gt 0) { $elapsedSec / $Completed } else { 0 }
    $etaSec = if ($Completed -gt 0) { $avg * ($Total - $Completed) } else { 0 }
    $rate = if ($elapsedSec -gt 0) { ($Completed / $elapsedSec) * 60.0 } else { 0 }
    $status = ('Repo {0}/{1}  |  Success {2}  Skip {3}  Fail {4}  |  elapsed {5}  ETA {6}  ({7:N1} repos/min)' -f `
        $Completed, $Total, $SuccessCount, $SkipCount, $FailCount,
        (Format-Duration $elapsedSec), (Format-Duration $etaSec), $rate)
    $current = if ($CurrentProject) { ('Project {0}/{1}: {2}' -f $ProjIndex, $ProjTotal, $CurrentProject) } else { '' }
    Write-Progress -Id $script:ProgressIdOverall `
        -Activity ('Mirror {0} -> {1}' -f $script:SourceCollection, $script:TargetCollection) `
        -Status $status `
        -PercentComplete $pct `
        -CurrentOperation $current
}

function Write-ProjectProgress {
    param(
        [string]$ProjectName,
        [int]$Completed,
        [int]$Total,
        [string]$CurrentRepo
    )
    if ($script:NoProgress) { return }
    $pct = if ($Total -gt 0) { [int](($Completed * 100) / $Total) } else { 0 }
    Write-Progress -Id $script:ProgressIdProject -ParentId $script:ProgressIdOverall `
        -Activity ('Project: {0}' -f $ProjectName) `
        -Status ('Repo {0}/{1}' -f ($Completed + 1), $Total) `
        -PercentComplete $pct `
        -CurrentOperation $CurrentRepo
}

function Write-StepProgress {
    param(
        [string]$RepoName,
        [string]$Step,           # short label, e.g. 'Cloning (mirror)'
        [int]$PercentComplete = -1
    )
    if ($script:NoProgress) { return }
    $params = @{
        Id       = $script:ProgressIdStep
        ParentId = $script:ProgressIdProject
        Activity = ('Repo: {0}' -f $RepoName)
        Status   = $Step
    }
    if ($PercentComplete -ge 0) { $params['PercentComplete'] = $PercentComplete }
    Write-Progress @params
}

function Clear-AllMirrorProgress {
    if ($script:NoProgress) { return }
    Write-Progress -Id $script:ProgressIdStep    -ParentId $script:ProgressIdProject -Activity ' ' -Completed
    Write-Progress -Id $script:ProgressIdProject -ParentId $script:ProgressIdOverall -Activity ' ' -Completed
    Write-Progress -Id $script:ProgressIdOverall                                      -Activity ' ' -Completed
}

# One-line scroll-back-friendly tick printed after every repo so you have a
# permanent text trail even when the live Write-Progress bars scroll away or
# the script is run with -NoProgress.
function Write-MirrorTick {
    param(
        [int]$Idx,
        [int]$Total,
        [PSCustomObject]$Item,
        [int]$Success,
        [int]$Skip,
        [int]$Fail,
        [System.Diagnostics.Stopwatch]$Sw
    )
    $pct = if ($Total -gt 0) { ($Idx * 100.0) / $Total } else { 0 }
    $elapsedSec = $Sw.Elapsed.TotalSeconds
    $avg = if ($Idx -gt 0) { $elapsedSec / $Idx } else { 0 }
    $etaSec = if ($Idx -gt 0) { $avg * ($Total - $Idx) } else { 0 }
    $symbol = switch ($Item.Status) {
        'Success'     { '[OK]'   }
        'DryRun'      { '[DRY]'  }
        'Skipped'     { '[SKIP]' }
        'Failed'      { '[FAIL]' }
        'TimedOut'    { '[TIME]' }
        'PathTooLong' { '[PATH]' }
        default       { '[--]'   }
    }
    $color = switch ($Item.Status) {
        'Success' { 'Green' }
        'DryRun'  { 'Yellow' }
        'Skipped' { 'DarkGray' }
        default   { 'Red' }
    }
    $line = ('  [{0,4}/{1}] {2,5:N1}%  {3}  {4}/{5}  ({6}s)  totals: OK {7}  SKIP {8}  FAIL {9}  elapsed {10}  ETA {11}' -f `
        $Idx, $Total, $pct, $symbol, $Item.SourceProject, $Item.SourceRepo,
        $Item.DurationSeconds, $Success, $Skip, $Fail,
        (Format-Duration $elapsedSec), (Format-Duration $etaSec))
    Write-Host $line -ForegroundColor $color
}

# ─── Mirror Loop ──────────────────────────────────────────────────────────────

$results = [System.Collections.Generic.List[object]]::new()
$successCount = 0
$skipCount = 0
$failCount = 0
$idx = 0

# Group work by project for project-level progress + per-project totals
$projGroups  = $workItems | Group-Object -Property SourceProject
$projOrder   = $projGroups | Sort-Object Name | Select-Object -ExpandProperty Name
$projTotals  = @{}
foreach ($g in $projGroups) {
    $projTotals[$g.Name] = [PSCustomObject]@{
        Project = $g.Name
        Total   = $g.Count
        Done    = 0
        Success = 0
        Skipped = 0
        Failed  = 0
        Seconds = 0
    }
}
$currentProject = $null
$projIndex      = 0
$overallSw      = [System.Diagnostics.Stopwatch]::StartNew()

# Ensure every target project exists once (so empty-project mirroring still creates them)
$allTargetProjects = @($workItems | Select-Object -ExpandProperty SourceProject -Unique) +
    @($projectsWithoutRepos | Where-Object { $_ })
$allTargetProjects = $allTargetProjects | Sort-Object -Unique

foreach ($projName in $allTargetProjects) {
    $srcMeta = $srcProjects | Where-Object { $_.name -eq $projName } | Select-Object -First 1
    try {
        Resolve-TargetProject -ProjectName $projName -Description $srcMeta.description | Out-Null
    }
    catch {
        Write-MigrationLog -Message "Failed to ensure target project '$projName': $($_.Exception.Message)" -LogFile $logFile -Level ERROR
    }
}

foreach ($item in $workItems) {
    $idx++
    $key = "$($item.SourceCollection)|$($item.SourceProject)|$($item.SourceRepo)"
    $item.Timestamp = (Get-Date -Format 'yyyy-MM-dd HH:mm:ss')

    # Detect project boundary for project-level progress + section header
    if ($item.SourceProject -ne $currentProject) {
        $currentProject = $item.SourceProject
        $projIndex++
        # Reset the per-step bar when entering a new project
        if (-not $NoProgress) {
            Write-Progress -Id $ProgressIdStep -ParentId $ProgressIdProject -Activity ' ' -Completed
        }
    }

    $pInfo = $projTotals[$currentProject]

    # Update overall + project bars BEFORE doing work
    Write-OverallProgress -Completed ($idx - 1) -Total $workItems.Count `
        -SuccessCount $successCount -SkipCount $skipCount -FailCount $failCount `
        -CurrentProject $currentProject -ProjIndex $projIndex -ProjTotal $projOrder.Count `
        -OverallStopwatch $overallSw
    Write-ProjectProgress -ProjectName $currentProject -Completed $pInfo.Done -Total $pInfo.Total -CurrentRepo $item.SourceRepo
    Write-StepProgress -RepoName $item.SourceRepo -Step 'Starting...' -PercentComplete 0

    Show-MenuHeader -Title ("[{0}/{1}] {2}/{3}" -f $idx, $workItems.Count, $item.SourceProject, $item.SourceRepo)

    if ($resumeStatus.ContainsKey($key) -and $resumeStatus[$key] -eq 'Success') {
        Write-Host "  Skipping (prior run succeeded)" -ForegroundColor DarkGray
        Write-StepProgress -RepoName $item.SourceRepo -Step 'Skipped (prior run succeeded)' -PercentComplete 100
        $item.Status = 'Skipped'
        $item.Reason = 'Prior run succeeded'
        $results.Add($item)
        $skipCount++
        $pInfo.Done++; $pInfo.Skipped++
        Write-MirrorTick -Idx $idx -Total $workItems.Count -Item $item -Success $successCount -Skip $skipCount -Fail $failCount -Sw $overallSw
        $results | Export-Csv -Path $manifestPath -NoTypeInformation -Encoding utf8 -Force
        continue
    }

    $sw = [System.Diagnostics.Stopwatch]::StartNew()

    try {
        Write-StepProgress -RepoName $item.SourceRepo -Step 'Ensuring target project...' -PercentComplete 10
        $targetProj = Resolve-TargetProject -ProjectName $item.SourceProject

        # Ensure target repo exists
        $tgtRepo = $null
        if (-not $DryRun) {
            Write-StepProgress -RepoName $item.SourceRepo -Step 'Ensuring target repo...' -PercentComplete 20
            $tgtRepo = Get-AdoGitRepository -ServerUrl $targetUrl -Collection $TargetCollection `
                -ProjectIdOrName $targetProj.id -RepoName $item.TargetRepo -Pat $targetPat
            if (-not $tgtRepo) {
                Write-Host "  Creating target repo '$($item.TargetRepo)' in '$($item.TargetProject)'..." -ForegroundColor Cyan
                Write-StepProgress -RepoName $item.SourceRepo -Step 'Creating target repo...' -PercentComplete 25
                try {
                    $tgtRepo = New-AdoGitRepository -ServerUrl $targetUrl -Collection $TargetCollection `
                        -ProjectId $targetProj.id -Name $item.TargetRepo -Pat $targetPat
                }
                catch {
                    if ($_.Exception.Message -match '409|already exists') {
                        $tgtRepo = Get-AdoGitRepository -ServerUrl $targetUrl -Collection $TargetCollection `
                            -ProjectIdOrName $targetProj.id -RepoName $item.TargetRepo -Pat $targetPat
                    }
                    else { throw }
                }
            }
            else {
                Write-Host "  Target repo already exists — will force `git push --mirror`." -ForegroundColor DarkYellow
            }
            $item.TargetRepoUrl = $tgtRepo.remoteUrl
        }
        else {
            Write-Host "  [DRY RUN] Would ensure target repo '$($item.TargetProject)/$($item.TargetRepo)' exists" -ForegroundColor Yellow
        }

        # Empty source repo? Nothing to clone/push — just count as success.
        if ($item.SourceSize -and [int64]$item.SourceSize -eq 0) {
            Write-Host "  Source repo is empty (size=0) — target repo ensured, nothing to push." -ForegroundColor DarkGray
            $item.Status = 'Success'
            $item.Reason = 'Empty source — no commits'
        }
        elseif ($DryRun) {
            Write-Host "  [DRY RUN] Would clone --mirror from source and push --mirror to target" -ForegroundColor Yellow
            $item.Status = 'DryRun'
            $item.Reason = 'DryRun — no clone/push performed'
        }
        else {
            # Mirror clone
            $localBare = Join-Path $WorkingDirectory ("{0}__{1}.git" -f $item.SourceProject, $item.SourceRepo)
            $localBare = $localBare -replace '[\\/:*?"<>|]', '_'
            $localBare = Join-Path (Split-Path $localBare -Parent) (Split-Path $localBare -Leaf)

            if (Test-Path $localBare) {
                Write-Host "  Removing stale cache: $localBare" -ForegroundColor DarkGray
                Remove-Item -Path $localBare -Recurse -Force -ErrorAction SilentlyContinue
            }

            $sourceAuthUrl = Add-PatToGitUrl -Url $item.SourceRepoUrl -Pat $sourcePat
            $targetAuthUrl = Add-PatToGitUrl -Url $item.TargetRepoUrl -Pat $targetPat

            Write-StepProgress -RepoName $item.SourceRepo -Step 'Cloning (mirror) from source...' -PercentComplete 35
            $cloneArgs = "clone --mirror `"$sourceAuthUrl`" `"$localBare`""
            Invoke-GitMirror -Arguments $cloneArgs -LogFile $logFile -StatusLabel 'Cloning (mirror)' `
                -TimeoutMinutes $TimeoutMinutes -StallTimeoutMinutes $StallTimeoutMinutes | Out-Null

            Write-StepProgress -RepoName $item.SourceRepo -Step 'Pushing (mirror) to target...' -PercentComplete 70
            $pushArgs = "push --mirror `"$targetAuthUrl`""
            Invoke-GitMirror -Arguments $pushArgs -WorkingDirectory $localBare -LogFile $logFile `
                -StatusLabel 'Pushing (mirror)' -TimeoutMinutes $TimeoutMinutes -StallTimeoutMinutes $StallTimeoutMinutes | Out-Null

            if (-not $KeepCache) {
                Write-StepProgress -RepoName $item.SourceRepo -Step 'Cleaning cache...' -PercentComplete 95
                Remove-Item -Path $localBare -Recurse -Force -ErrorAction SilentlyContinue
            }

            $item.Status = 'Success'
            $item.Reason = ''
        }

        if ($item.Status -eq 'Success' -or $item.Status -eq 'DryRun') { $successCount++ }
    }
    catch {
        $msg = $_.Exception.Message
        $friendly = Get-FriendlyError -ErrorMessage $msg
        $item.Status = if ($msg -match 'too long|PathTooLong|260 char') { 'PathTooLong' }
                       elseif ($msg -match 'TIMED OUT|timed out|STALL') { 'TimedOut' }
                       else { 'Failed' }
        $item.Reason = $friendly
        Write-MigrationLog -Message "FAILED $($item.SourceProject)/$($item.SourceRepo): $msg" -LogFile $logFile -Level ERROR
        Write-Host "  ✗ $($item.Status): $friendly" -ForegroundColor Red
        $failCount++
    }
    finally {
        $sw.Stop()
        $item.DurationSeconds = [int]$sw.Elapsed.TotalSeconds
        $results.Add($item)

        # Per-project rollup
        $pInfo.Done++
        $pInfo.Seconds += $item.DurationSeconds
        switch ($item.Status) {
            { $_ -in @('Success', 'DryRun') } { $pInfo.Success++ }
            'Skipped'                         { $pInfo.Skipped++ }
            default                           { $pInfo.Failed++ }
        }

        # Mark this repo's step bar complete (project + overall bars update on next iteration)
        Write-StepProgress -RepoName $item.SourceRepo -Step ('Done — {0}' -f $item.Status) -PercentComplete 100

        # Scrollback tick — one line per repo, survives -NoProgress and terminal scroll
        Write-MirrorTick -Idx $idx -Total $workItems.Count -Item $item `
            -Success $successCount -Skip $skipCount -Fail $failCount -Sw $overallSw

        # Flush manifest after every item so a crash doesn't lose progress
        $results | Export-Csv -Path $manifestPath -NoTypeInformation -Encoding utf8 -Force
    }
}

# Dismiss the progress bars before printing the final summary
Clear-AllMirrorProgress

# ─── Final Report ─────────────────────────────────────────────────────────────

$reportPath = Join-Path $config.outputDirectory "mirror-REPORT-$timestamp.json"
$report = [ordered]@{
    timestamp        = (Get-Date -Format 'o')
    sourceCollection = $SourceCollection
    sourceUrl        = $sourceUrl
    targetCollection = $TargetCollection
    targetUrl        = $targetUrl
    dryRun           = [bool]$DryRun
    totals           = [ordered]@{
        projects = $srcProjects.Count
        repos    = $workItems.Count
        success  = $successCount
        skipped  = $skipCount
        failed   = $failCount
    }
    manifestPath     = $manifestPath
    previewPath      = $previewPath
    logFile          = $logFile
}
$report | ConvertTo-Json -Depth 6 | Out-File -FilePath $reportPath -Encoding utf8

Write-Host ""
Show-MenuHeader -Title 'Mirror complete'
Write-Host ("  Total elapsed      : {0}" -f (Format-Duration $overallSw.Elapsed.TotalSeconds)) -ForegroundColor White
Write-Host "  Projects processed : $($srcProjects.Count)" -ForegroundColor White
Write-Host "  Repos processed    : $($workItems.Count)" -ForegroundColor White
Write-Host "  Successful         : $successCount" -ForegroundColor Green
Write-Host "  Skipped (resume)   : $skipCount" -ForegroundColor DarkGray
Write-Host "  Failed             : $failCount" -ForegroundColor $(if ($failCount -gt 0) { 'Red' } else { 'DarkGray' })

# ─── Per-Project Roll-Up ──────────────────────────────────────────────────────
if ($projTotals.Count -gt 0) {
    Write-Host ""
    Write-Host "Per-project results:" -ForegroundColor Cyan
    $rows = foreach ($name in ($projTotals.Keys | Sort-Object)) {
        $p = $projTotals[$name]
        [PSCustomObject]@{
            Project = $p.Project
            Repos   = $p.Total
            Success = $p.Success
            Skipped = $p.Skipped
            Failed  = $p.Failed
            Elapsed = (Format-Duration $p.Seconds)
        }
    }
    $rows | Format-Table -AutoSize | Out-String | Write-Host
}

Write-Host ""
Write-Host "  Manifest CSV: $manifestPath" -ForegroundColor Cyan
Write-Host "  Report JSON : $reportPath" -ForegroundColor Cyan
Write-Host "  Log         : $logFile" -ForegroundColor Cyan
if ($failCount -gt 0) {
    Write-Host ""
    Write-Host "Some repos failed. To retry just those, re-run with:" -ForegroundColor Yellow
    Write-Host "  -ResumeManifest `"$manifestPath`"" -ForegroundColor Yellow
    exit 1
}
