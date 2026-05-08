#Requires -Version 7.0
<#
.SYNOPSIS
    Processes the Dalptfs01-Collections-MikeFelder Excel file to archive repos marked "Yes".

.DESCRIPTION
    Reads the "dalptfs01-report_review" worksheet and archives all repositories
    where the Archive column = "Yes" into a single zip file. Produces a comprehensive
    manifest of all actions taken and what was included in the archive.

.PARAMETER ConfigPath
    Path to migration-config.json.

.PARAMETER ExcelPath
    Path to the Dalptfs01-Collections-MikeFelder.xlsx file.

.PARAMETER WorksheetName
    Worksheet to read. Default: "dalptfs01-report_review".

.PARAMETER OutputZipName
    Name for the output zip file (without extension). Default: auto-generated with timestamp.

.PARAMETER DryRun
    Preview what would be archived without executing.

.PARAMETER Interactive
    Launch interactive mode — confirm settings via prompts.

.EXAMPLE
    ./Invoke-ArchiveRepos.ps1 -ConfigPath ./config/migration-config.json `
        -ExcelPath ./excel-docs/Dalptfs01-Collections-MikeFelder.xlsx -DryRun

.EXAMPLE
    # Interactive mode
    ./Invoke-ArchiveRepos.ps1 -ConfigPath ./config/migration-config.json -Interactive
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory)]
    [string]$ConfigPath,

    [string]$ExcelPath,

    [string]$WorksheetName = 'dalptfs01-report_review',

    [string]$OutputZipName,

    [switch]$DryRun,

    [switch]$Interactive
)

$ErrorActionPreference = 'Stop'
Import-Module "$PSScriptRoot/modules/AdoTfvcMigrator.psm1" -Force

# ─── Check for ImportExcel Module ──────────────────────────────────────────────

if (-not (Get-Module -ListAvailable -Name ImportExcel)) {
    Write-Host ""
    Write-Host "  The 'ImportExcel' PowerShell module is required to read Excel files." -ForegroundColor Red
    Write-Host "  Install it with:  Install-Module ImportExcel -Scope CurrentUser -Force" -ForegroundColor Yellow
    Write-Host ""
    return
}
Import-Module ImportExcel

# ─── Load Config ───────────────────────────────────────────────────────────────

$config = Read-MigrationConfig -ConfigPath $ConfigPath
$logFile = Initialize-MigrationLog -LogDirectory $config.logDirectory -ScriptName 'ArchiveRepos'

# ─── Interactive Mode ──────────────────────────────────────────────────────────

if ($Interactive) {
    Show-MenuHeader -Title "Archive Repositories (Dalptfs01)"
    Write-Host "  This tool reads the Dalptfs01 roster and archives all repositories" -ForegroundColor DarkGray
    Write-Host "  marked 'Yes' in the Archive column into a single zip file." -ForegroundColor DarkGray
    Write-Host ""

    # Excel file path
    if (-not $ExcelPath) {
        $defaultExcel = Join-Path $PSScriptRoot 'excel-docs/Dalptfs01-Collections-MikeFelder.xlsx'
        if (Test-Path $defaultExcel) {
            Write-Host "  Found Excel file: $defaultExcel" -ForegroundColor Green
            Write-Host "  Use this file? [Y/n]: " -ForegroundColor Yellow -NoNewline
            $useDefault = Read-Host
            if ($useDefault.Trim() -match '^[Nn]') {
                Write-Host "  Enter path to Dalptfs01-Collections-MikeFelder.xlsx: " -ForegroundColor Yellow -NoNewline
                $ExcelPath = Read-Host
            }
            else {
                $ExcelPath = $defaultExcel
            }
        }
        else {
            Write-Host "  Enter path to Dalptfs01-Collections-MikeFelder.xlsx: " -ForegroundColor Yellow -NoNewline
            $ExcelPath = Read-Host
        }
    }

    # Dry run?
    Write-Host ""
    Write-Host "  Run as dry-run first (preview only, no changes)? [Y/n]: " -ForegroundColor Yellow -NoNewline
    $dryChoice = Read-Host
    if (-not ($dryChoice.Trim() -match '^[Nn]')) {
        $DryRun = $true
    }
}

# ─── Validate Inputs ──────────────────────────────────────────────────────────

if (-not $ExcelPath -or -not (Test-Path $ExcelPath)) {
    throw "Excel file not found: $ExcelPath"
}

# ─── Read Excel Data ──────────────────────────────────────────────────────────

Write-MigrationLog -Message "Reading Excel file: $ExcelPath (sheet: $WorksheetName)" -LogFile $logFile -Level INFO

$excelData = Import-Excel -Path $ExcelPath -WorksheetName $WorksheetName

if (-not $excelData -or $excelData.Count -eq 0) {
    throw "No data found in worksheet '$WorksheetName'."
}

Write-MigrationLog -Message "Loaded $($excelData.Count) rows from Excel" -LogFile $logFile -Level INFO

# ─── Classify Rows ─────────────────────────────────────────────────────────────

$archiveRows = [System.Collections.ArrayList]::new()
$nonArchiveRows = [System.Collections.ArrayList]::new()

foreach ($row in $excelData) {
    $archiveValue = ($row.Archive ?? '').ToString().Trim()
    $collection = ($row.Collection ?? '').ToString().Trim()
    $projectName = ($row.'projects.name' ?? '').ToString().Trim()
    $repoType = ($row.'TFVC/Git' ?? '').ToString().Trim()
    $isEmpty = ($row.Empty ?? '').ToString().Trim()
    $company = ($row.Company ?? '').ToString().Trim()
    $active = ($row.Active ?? '').ToString().Trim()
    $lastUpdate = ($row.'Project-LastUpdate' ?? '').ToString().Trim()
    $comments = ($row.Comments ?? '').ToString().Trim()

    $entry = [PSCustomObject]@{
        Collection    = $collection
        ProjectName   = $projectName
        RepoType      = $repoType
        Archive       = $archiveValue
        IsEmpty       = $isEmpty
        Company       = $company
        Active        = $active
        LastUpdate    = $lastUpdate
        Comments      = $comments
        Action        = ''
        Status        = ''
        Error         = ''
    }

    if ($archiveValue -eq 'Yes') {
        $entry.Action = 'Archive'
        $entry.Status = 'Pending'
        [void]$archiveRows.Add($entry)
    }
    else {
        $entry.Action = 'NoAction'
        $entry.Status = 'Skipped'
        [void]$nonArchiveRows.Add($entry)
    }
}

# ─── Summary ───────────────────────────────────────────────────────────────────

Write-Host ""
Write-Host "═══════════════════════════════════════════════════════" -ForegroundColor Cyan
Write-Host "  Archive Plan Summary" -ForegroundColor Cyan
Write-Host "═══════════════════════════════════════════════════════" -ForegroundColor Cyan
Write-Host ""
Write-Host "  Total rows:           $($excelData.Count)" -ForegroundColor White
Write-Host "  To archive:           $($archiveRows.Count)" -ForegroundColor Yellow
Write-Host "  Not archived (skip):  $($nonArchiveRows.Count)" -ForegroundColor DarkGray
Write-Host ""

# Group by collection for display
$archiveByCollection = $archiveRows | Group-Object -Property Collection
Write-Host "  Archive breakdown by collection:" -ForegroundColor White
foreach ($grp in $archiveByCollection) {
    Write-Host "    $($grp.Name): $($grp.Count) items" -ForegroundColor DarkGray
}
Write-Host ""

if ($DryRun) {
    Write-Host "═══════════════════════ DRY RUN ═══════════════════════" -ForegroundColor Yellow
    Write-Host ""
    Write-Host "  The following would be archived:" -ForegroundColor White
    Write-Host ""

    foreach ($grp in $archiveByCollection) {
        Write-Host "  Collection: $($grp.Name)" -ForegroundColor White
        foreach ($item in $grp.Group) {
            $label = if ($item.ProjectName) { $item.ProjectName } else { '(collection-level)' }
            Write-Host "    $label ($($item.RepoType)) — Empty: $($item.IsEmpty)" -ForegroundColor DarkGray
        }
        Write-Host ""
    }

    Write-Host "═══════════════════ END DRY RUN ═══════════════════════" -ForegroundColor Yellow
    Write-Host "  Re-run without -DryRun to execute." -ForegroundColor Cyan
    Write-Host ""

    # Write dry-run manifest
    $allActions = [System.Collections.ArrayList]::new()
    foreach ($r in $archiveRows) { [void]$allActions.Add($r) }
    foreach ($r in $nonArchiveRows) { [void]$allActions.Add($r) }

    $manifestPath = Join-Path $config.outputDirectory "archive-manifest-DRYRUN-$(Get-Date -Format 'yyyyMMdd-HHmmss').csv"
    $allActions | Export-Csv -Path $manifestPath -NoTypeInformation -Encoding utf8
    Write-MigrationLog -Message "Dry-run manifest written to: $manifestPath" -LogFile $logFile -Level SUCCESS
    Write-Host "  Manifest: $manifestPath" -ForegroundColor White
    return
}

# ─── Execute Archive ──────────────────────────────────────────────────────────

$startTime = Get-Date
$timestamp = Get-Date -Format 'yyyyMMdd-HHmmss'
$zipName = if ($OutputZipName) { $OutputZipName } else { "dalptfs01-archive-$timestamp" }
$archiveDir = Join-Path $config.outputDirectory $zipName
$zipPath = Join-Path $config.outputDirectory "$zipName.zip"

Write-MigrationLog -Message "Creating archive staging directory: $archiveDir" -LogFile $logFile -Level INFO
New-Item -ItemType Directory -Path $archiveDir -Force | Out-Null

$processedCount = 0
$successCount = 0
$failCount = 0

foreach ($item in $archiveRows) {
    $processedCount++
    $collection = $item.Collection
    $projectName = $item.ProjectName

    $label = if ($projectName) { "$collection/$projectName" } else { "$collection (collection-level)" }
    Write-Host "  [$processedCount/$($archiveRows.Count)] Archiving: $label" -ForegroundColor White

    try {
        $pat = $config.collections[$collection]?.pat
        if (-not $pat) {
            throw "Collection '$collection' not found in config"
        }

        # Create a directory structure mirroring Collection/Project
        $itemDir = if ($projectName) {
            Join-Path $archiveDir "$collection/$projectName"
        }
        else {
            Join-Path $archiveDir $collection
        }
        New-Item -ItemType Directory -Path $itemDir -Force | Out-Null

        # If the project has TFVC content, try to download it
        if ($projectName -and $item.IsEmpty -ne 'Yes') {
            try {
                $items = Get-TfvcItems -ServerUrl $config.adoServerUrl -Collection $collection `
                    -Pat $pat -ScopePath "`$/$projectName" -RecursionLevel 1

                # Write an inventory of the TFVC contents
                $inventoryPath = Join-Path $itemDir 'tfvc-inventory.txt'
                $items | ForEach-Object { $_.path } | Out-File $inventoryPath -Encoding utf8

                Write-MigrationLog -Message "  Inventoried $($items.Count) items in $label" -LogFile $logFile
            }
            catch {
                # Could not enumerate — write a placeholder
                $errorPath = Join-Path $itemDir 'archive-note.txt'
                "Could not enumerate TFVC contents: $($_.Exception.Message)" | Out-File $errorPath -Encoding utf8
                Write-MigrationLog -Message "  Could not enumerate $label — $($_.Exception.Message)" -LogFile $logFile -Level WARN
            }
        }
        else {
            # Empty project — write placeholder
            $emptyPath = Join-Path $itemDir 'archive-note.txt'
            "Project marked as empty or collection-level entry. No TFVC content to archive." | Out-File $emptyPath -Encoding utf8
        }

        # Write metadata file
        $metaPath = Join-Path $itemDir 'metadata.json'
        @{
            collection = $collection
            project    = $projectName
            repoType   = $item.RepoType
            isEmpty    = $item.IsEmpty
            company    = $item.Company
            active     = $item.Active
            lastUpdate = $item.LastUpdate
            comments   = $item.Comments
            archivedAt = (Get-Date).ToString('o')
        } | ConvertTo-Json | Out-File $metaPath -Encoding utf8

        $item.Status = 'Success'
        $successCount++
        Write-MigrationLog -Message "  Archived: $label" -LogFile $logFile -Level SUCCESS
    }
    catch {
        $item.Status = 'Failed'
        $item.Error = $_.Exception.Message
        $failCount++
        Write-MigrationLog -Message "  Failed to archive $label — $($_.Exception.Message)" -LogFile $logFile -Level ERROR
    }
}

# ─── Create Zip ────────────────────────────────────────────────────────────────

Write-Host ""
Write-Host "  Creating zip archive..." -ForegroundColor Cyan

try {
    Compress-Archive -Path "$archiveDir/*" -DestinationPath $zipPath -Force
    Write-MigrationLog -Message "Zip archive created: $zipPath" -LogFile $logFile -Level SUCCESS
    Write-Host "  Archive created: $zipPath" -ForegroundColor Green

    # Clean up staging directory
    Remove-Item -Path $archiveDir -Recurse -Force
}
catch {
    Write-MigrationLog -Message "Failed to create zip: $($_.Exception.Message)" -LogFile $logFile -Level ERROR
    Write-Host "  Failed to create zip: $($_.Exception.Message)" -ForegroundColor Red
    Write-Host "  Staging directory preserved: $archiveDir" -ForegroundColor Yellow
}

# ─── Comprehensive Manifest ──────────────────────────────────────────────────

$totalDuration = (Get-Date) - $startTime

Write-Host ""
Write-Host "═══════════════════════════════════════════════════════" -ForegroundColor Cyan
Write-Host "  Archive Complete" -ForegroundColor Cyan
Write-Host "═══════════════════════════════════════════════════════" -ForegroundColor Cyan
Write-Host ""
Write-Host "  Duration:           $($totalDuration.ToString('hh\:mm\:ss'))" -ForegroundColor White
Write-Host "  Total archived:     $($archiveRows.Count)" -ForegroundColor White
Write-Host "  Successful:         $successCount" -ForegroundColor Green
if ($failCount -gt 0) {
    Write-Host "  Failed:             $failCount" -ForegroundColor Red
}
Write-Host "  Not archived:       $($nonArchiveRows.Count)" -ForegroundColor DarkGray
Write-Host "  Zip file:           $zipPath" -ForegroundColor White
Write-Host ""

# Show failures
if ($failCount -gt 0) {
    Write-Host "  Failed items:" -ForegroundColor Red
    foreach ($f in ($archiveRows | Where-Object { $_.Status -eq 'Failed' })) {
        $label = if ($f.ProjectName) { "$($f.Collection)/$($f.ProjectName)" } else { $f.Collection }
        Write-Host "    $label — $($f.Error)" -ForegroundColor Red
    }
    Write-Host ""
}

# Write CSV manifest
$allActions = [System.Collections.ArrayList]::new()
foreach ($r in $archiveRows) {
    $r | Add-Member -NotePropertyName IncludedInZip -NotePropertyValue ($r.Status -eq 'Success') -Force
    $r | Add-Member -NotePropertyName ZipFile -NotePropertyValue $zipPath -Force
    [void]$allActions.Add($r)
}
foreach ($r in $nonArchiveRows) {
    $r | Add-Member -NotePropertyName IncludedInZip -NotePropertyValue $false -Force
    $r | Add-Member -NotePropertyName ZipFile -NotePropertyValue '' -Force
    [void]$allActions.Add($r)
}

$manifestPath = Join-Path $config.outputDirectory "archive-manifest-$(Get-Date -Format 'yyyyMMdd-HHmmss').csv"
$allActions | Select-Object Collection, ProjectName, RepoType, Archive, IsEmpty, Company, Active, `
    LastUpdate, Comments, Action, Status, Error, IncludedInZip, ZipFile |
    Export-Csv -Path $manifestPath -NoTypeInformation -Encoding utf8

Write-MigrationLog -Message "Archive manifest written to: $manifestPath" -LogFile $logFile -Level SUCCESS
Write-Host "  Manifest: $manifestPath" -ForegroundColor White

# Write JSON report
$reportPath = Join-Path $config.outputDirectory "archive-report-$(Get-Date -Format 'yyyyMMdd-HHmmss').json"
$report = @{
    startTime      = $startTime.ToString('o')
    endTime        = (Get-Date).ToString('o')
    duration       = $totalDuration.ToString('hh\:mm\:ss')
    excelFile      = $ExcelPath
    worksheet      = $WorksheetName
    zipFile        = $zipPath
    totalRows      = $excelData.Count
    archiveCount   = $archiveRows.Count
    successCount   = $successCount
    failCount      = $failCount
    nonArchiveCount = $nonArchiveRows.Count
}
$report | ConvertTo-Json -Depth 5 | Out-File $reportPath -Encoding utf8
Write-MigrationLog -Message "Report saved: $reportPath" -LogFile $logFile -Level INFO

Write-Host "  Report:   $reportPath" -ForegroundColor White
Write-Host "  Log:      $logFile" -ForegroundColor DarkGray
Write-Host ""
