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

.PARAMETER ConfigPath
    Path to migration-config.json.

.PARAMETER Collections
    Optional. Limit discovery to specific collection names. Default: all collections in config.

.PARAMETER Depth
    How many folder levels deep to enumerate. Default: 2.

.PARAMETER GenerateAuthorMap
    If set, generates an author-mapping CSV template from discovered identities.

.PARAMETER OutputFormat
    Output format: Table, Json, Csv. Default: Table.

.EXAMPLE
    ./Invoke-TfvcDiscovery.ps1 -ConfigPath ./config/migration-config.json

.EXAMPLE
    ./Invoke-TfvcDiscovery.ps1 -ConfigPath ./config/migration-config.json -Collections GAMS -Depth 3 -OutputFormat Json
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory)]
    [string]$ConfigPath,

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

foreach ($collectionName in $targetCollections) {
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

    foreach ($project in $projects) {
        $projectName = $project.name
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
        $folders = $rootItems | Where-Object { $_.isFolder -eq $true -and $_.path -ne "`$/$projectName" }

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

            $entry = [PSCustomObject]@{
                Collection     = $collectionName
                Project        = $projectName
                TfvcPath       = $folderPath
                SubFolders     = $subItemCount
                LastChangeDate = $lastDate
                LastAuthor     = $lastAuthor
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
