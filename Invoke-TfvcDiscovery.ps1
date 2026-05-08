#Requires -Version 7.0
<#
.SYNOPSIS
    Discovers and inventories all TFVC repositories and folder structures across ADO collections.

.DESCRIPTION
    Connects to one or more ADO 2022 collections and enumerates:
    - All team projects with TFVC repos
    - Top-level folder structure under each $/Project
    - Branch information
    - Last changeset per path (activity indicator)
    - Unique committer identities (for author mapping)

    Supports -Interactive mode to select collections and output format via menus.

.PARAMETER ConfigPath
    Path to migration-config.json.

.PARAMETER Interactive
    Launch interactive mode — select collections and options via menus.

.PARAMETER Collections
    Optional. Limit discovery to specific collection names. Default: all collections in config.

.PARAMETER Depth
    How many folder levels deep to enumerate. Default: 2.

.PARAMETER GenerateAuthorMap
    If set, generates an author-mapping CSV template from discovered identities.

.PARAMETER OutputFormat
    Output format: Table, Json, Csv. Default: Table.

.EXAMPLE
    # Interactive mode
    ./Invoke-TfvcDiscovery.ps1 -ConfigPath ./config/migration-config.json -Interactive

.EXAMPLE
    ./Invoke-TfvcDiscovery.ps1 -ConfigPath ./config/migration-config.json -Collections GAMS -Depth 3 -OutputFormat Json
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory)]
    [string]$ConfigPath,

    [switch]$Interactive,

    [string[]]$Collections,

    [int]$Depth = 2,

    [switch]$GenerateAuthorMap,

    [ValidateSet('Table', 'Json', 'Csv')]
    [string]$OutputFormat = 'Table'
)

$ErrorActionPreference = 'Stop'
Import-Module "$PSScriptRoot/modules/AdoTfvcMigrator.psm1" -Force

# ─── Load Config ───────────────────────────────────────────────────────────────

$config = Read-MigrationConfig -ConfigPath $ConfigPath
$logFile = Initialize-MigrationLog -LogDirectory $config.logDirectory -ScriptName 'TfvcDiscovery'

# ─── Interactive Mode ──────────────────────────────────────────────────────────

if ($Interactive) {
    Show-MenuHeader -Title "Discover TFVC Repositories"
    Write-Host "This will scan your ADO collections and list all TFVC repos and folders." -ForegroundColor DarkGray
    Write-Host "No changes are made — this is a read-only scan." -ForegroundColor DarkGray

    # 1. Pick collections
    $collectionNames = @($config.collections.Keys | Sort-Object)

    Write-Host ""
    Write-Host "  Which collections to scan?" -ForegroundColor White
    Write-Host ""
    Write-Host "  [1] All collections ($($collectionNames.Count) configured)" -ForegroundColor White
    Write-Host "  [2] Pick specific collection(s)" -ForegroundColor White
    Write-Host "  [0] Cancel" -ForegroundColor DarkGray
    Write-Host ""
    Write-Host "  Choice: " -ForegroundColor Yellow -NoNewline
    $collChoice = Read-Host

    switch ($collChoice.Trim()) {
        '0' { Write-Host "Cancelled." -ForegroundColor Yellow; return }
        '2' {
            $displayItems = $collectionNames | ForEach-Object {
                $desc = $config.collections[$_].description
                if ($desc) { "$_ — $desc" } else { $_ }
            }
            Show-MenuHeader -Title "Select collection(s) to scan"
            $selectedIndices = Show-NumberedMenu -Items $displayItems -Prompt 'Select collection(s)' -MultiSelect -AllowBack
            if ($null -eq $selectedIndices) { Write-Host "Cancelled." -ForegroundColor Yellow; return }
            $Collections = @($selectedIndices | ForEach-Object { $collectionNames[$_] })
        }
    }

    # 2. Depth
    Write-Host ""
    Write-Host "  How many folder levels deep to scan? [2]: " -ForegroundColor Yellow -NoNewline
    $depthInput = Read-Host
    if ($depthInput.Trim() -match '^\d+$') {
        $Depth = [int]$depthInput.Trim()
    }

    # 3. Output format
    Write-Host ""
    Write-Host "  Output format:" -ForegroundColor White
    Write-Host "  [1] Table (display on screen)" -ForegroundColor White
    Write-Host "  [2] CSV file (for Excel)" -ForegroundColor White
    Write-Host "  [3] JSON file" -ForegroundColor White
    Write-Host ""
    Write-Host "  Choice [1]: " -ForegroundColor Yellow -NoNewline
    $fmtChoice = Read-Host

    switch ($fmtChoice.Trim()) {
        '2' { $OutputFormat = 'Csv' }
        '3' { $OutputFormat = 'Json' }
        default { $OutputFormat = 'Table' }
    }

    # 4. Author map
    Write-Host ""
    Write-Host "  Generate an author mapping template (for converting usernames to Git emails)? [y/N]: " -ForegroundColor Yellow -NoNewline
    $authorChoice = Read-Host
    if ($authorChoice.Trim() -match '^[Yy]') {
        $GenerateAuthorMap = $true
    }

    Write-Host ""
    Write-Host "  Starting scan..." -ForegroundColor Cyan
    Write-Host ""
}

Write-MigrationLog -Message "Starting TFVC discovery" -LogFile $logFile -Level INFO
Write-MigrationLog -Message "ADO Server: $($config.adoServerUrl)" -LogFile $logFile

# ─── Determine Collections ────────────────────────────────────────────────────

$targetCollections = if ($Collections) {
    $Collections
} else {
    $config.collections.Keys
}

# ─── Discovery ─────────────────────────────────────────────────────────────────

$allResults = [System.Collections.ArrayList]::new()
$allAuthors = [System.Collections.ArrayList]::new()
$collectionIndex = 0
$collectionTotal = @($targetCollections).Count

foreach ($collectionName in $targetCollections) {
    $collectionIndex++
    Write-Host "  Scanning collection $collectionIndex of ${collectionTotal}: $collectionName" -ForegroundColor Cyan

    $collectionConfig = $config.collections[$collectionName]
    if (-not $collectionConfig) {
        Write-MigrationLog -Message "Collection '$collectionName' not found in config — skipping" -LogFile $logFile -Level WARN
        continue
    }

    $pat = $collectionConfig.pat
    Write-MigrationLog -Message "Connecting to collection: $collectionName" -LogFile $logFile

    # Test connection
    $connTest = Test-AdoConnection -ServerUrl $config.adoServerUrl -Collection $collectionName -Pat $pat
    if (-not $connTest.Connected) {
        Write-MigrationLog -Message "Cannot connect to $collectionName — $($connTest.Error)" -LogFile $logFile -Level ERROR
        continue
    }

    Write-MigrationLog -Message "Connected. Found $($connTest.ProjectCount) projects." -LogFile $logFile -Level SUCCESS

    # Get all projects
    $projects = Get-AdoProjects -ServerUrl $config.adoServerUrl -Collection $collectionName -Pat $pat
    $projectIndex = 0
    $projectTotal = @($projects).Count

    foreach ($project in $projects) {
        $projectIndex++
        $projectName = $project.name
        Write-Host "    Project $projectIndex of ${projectTotal}: $projectName" -ForegroundColor DarkGray
        Write-MigrationLog -Message "  Scanning project: $projectName" -LogFile $logFile

        # Check for TFVC repo
        try {
            $rootItems = Get-TfvcItems -ServerUrl $config.adoServerUrl -Collection $collectionName `
                -Pat $pat -ScopePath "`$/$projectName" -RecursionLevel 1
        }
        catch {
            Write-MigrationLog -Message "  No TFVC content in $projectName (or access denied)" -LogFile $logFile -Level WARN
            continue
        }

        if (-not $rootItems -or $rootItems.Count -eq 0) {
            continue
        }

        # Get folders at the requested depth
        $folders = $rootItems | Where-Object {
            $_.isFolder -eq $true -and
            $_.path -ne "`$/$projectName" -and
            $_.path -notlike '*/BuildProcessTemplates'
        }

        # Fetch build definitions for this project to check pipeline associations
        $buildDefs = @()
        try {
            $buildDefs = @(Get-AdoBuildDefinitions -ServerUrl $config.adoServerUrl `
                -Collection $collectionName -Pat $pat -ProjectName $projectName)
        }
        catch { }

        foreach ($folder in $folders) {
            $folderPath = $folder.path

            # Get last changeset for activity info
            $lastChangeset = $null
            try {
                $lastChangeset = Get-TfvcChangesets -ServerUrl $config.adoServerUrl `
                    -Collection $collectionName -Pat $pat -ScopePath $folderPath -Top 1
            }
            catch {
                # Some paths may not have changesets
            }

            $lastDate = if ($lastChangeset) { $lastChangeset[0].createdDate } else { 'Unknown' }
            $lastAuthor = if ($lastChangeset) { $lastChangeset[0].author.displayName } else { 'Unknown' }
            $lastAuthorId = if ($lastChangeset) { $lastChangeset[0].author.uniqueName } else { $null }

            if ($lastAuthorId) {
                [void]$allAuthors.Add($lastAuthorId)
            }

            # Count sub-items if going deeper
            $subItemCount = 0
            if ($Depth -gt 1) {
                try {
                    $subItems = Get-TfvcItems -ServerUrl $config.adoServerUrl `
                        -Collection $collectionName -Pat $pat -ScopePath $folderPath -RecursionLevel 1
                    $subItemCount = ($subItems | Where-Object { $_.isFolder }).Count
                }
                catch { }
            }

            # Check if any pipeline/build definition references this folder path
            $hasPipeline = $false
            foreach ($bd in $buildDefs) {
                $repoPath = $bd.repository.defaultBranch ?? $bd.repository.rootFolder ?? ''
                if ($repoPath -like "*$($folder.path)*" -or $repoPath -like "*$projectName*") {
                    $hasPipeline = $true
                    break
                }
            }
            # Also check by name match (build def name often matches folder name)
            if (-not $hasPipeline) {
                $folderName = ($folderPath -split '/')[-1]
                foreach ($bd in $buildDefs) {
                    if ($bd.name -like "*$folderName*" -or ($bd.repository.name -and $bd.repository.name -eq $folderName)) {
                        $hasPipeline = $true
                        break
                    }
                }
            }

            $entry = [PSCustomObject]@{
                Collection     = $collectionName
                Project        = $projectName
                TfvcPath       = $folderPath
                SubFolders     = $subItemCount
                LastChangeDate = $lastDate
                LastAuthor     = $lastAuthor
                HasPipeline    = $hasPipeline
            }

            [void]$allResults.Add($entry)
        }

        # Enumerate branches for this project
        try {
            $branches = Get-TfvcBranches -ServerUrl $config.adoServerUrl -Collection $collectionName -Pat $pat
            $projectBranches = $branches | Where-Object { $_.path -like "`$/$projectName/*" }
            if ($projectBranches) {
                Write-MigrationLog -Message "  Branches found: $($projectBranches.Count)" -LogFile $logFile
                foreach ($branch in $projectBranches) {
                    Write-MigrationLog -Message "    $($branch.path)" -LogFile $logFile
                }
            }
        }
        catch {
            # Branch enumeration is informational
        }
    }
}

# ─── Output ────────────────────────────────────────────────────────────────────

Write-MigrationLog -Message "Discovery complete. Found $($allResults.Count) TFVC paths." -LogFile $logFile -Level SUCCESS

$outputPath = Join-Path $config.outputDirectory "tfvc-inventory.$(Get-Date -Format 'yyyyMMdd-HHmmss')"

switch ($OutputFormat) {
    'Json' {
        $jsonPath = "${outputPath}.json"
        $allResults | ConvertTo-Json -Depth 5 | Out-File $jsonPath -Encoding utf8
        Write-MigrationLog -Message "Inventory written to: $jsonPath" -LogFile $logFile -Level SUCCESS
    }
    'Csv' {
        $csvPath = "${outputPath}.csv"
        $allResults | Export-Csv -Path $csvPath -NoTypeInformation -Encoding utf8
        Write-MigrationLog -Message "Inventory written to: $csvPath" -LogFile $logFile -Level SUCCESS
    }
    default {
        $allResults | Format-Table -AutoSize
    }
}

# ─── Author Map Generation ────────────────────────────────────────────────────

if ($GenerateAuthorMap -and $allAuthors.Count -gt 0) {
    $authorMapPath = Join-Path $config.outputDirectory "author-mapping-template.csv"
    Export-AuthorMappingTemplate -TfvcIdentities $allAuthors -OutputPath $authorMapPath
    Write-MigrationLog -Message "Author mapping template written to: $authorMapPath" -LogFile $logFile -Level SUCCESS
    Write-Host ""
    Write-Host "Fill in GitName and GitEmail columns, then set authorMappingFile in your config." -ForegroundColor Cyan
}

Write-MigrationLog -Message "Log file: $logFile" -LogFile $logFile -Level INFO
