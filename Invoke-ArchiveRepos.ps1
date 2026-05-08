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

.PARAMETER Interactive
    Launch interactive mode — step-by-step guided flow with preview and confirmation.

.EXAMPLE
    # Interactive mode (recommended)
    ./Invoke-ArchiveRepos.ps1 -ConfigPath ./config/migration-config.json -Interactive

.EXAMPLE
    ./Invoke-ArchiveRepos.ps1 -ConfigPath ./config/migration-config.json `
        -ExcelPath ./excel-docs/Dalptfs01-Collections-MikeFelder.xlsx
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory)]
    [string]$ConfigPath,

    [string]$ExcelPath,

    [string]$WorksheetName = 'dalptfs01-report_review',

    [string]$OutputZipName,

    [switch]$Interactive
)

$ErrorActionPreference = 'Stop'
Import-Module "$PSScriptRoot/modules/AdoTfvcMigrator.psm1" -Force

# ─── Check for ImportExcel Module ──────────────────────────────────────────────

if (-not (Get-Module -ListAvailable -Name ImportExcel)) {
    Write-Host ""
    Write-Host "  ┌──────────────────────────────────────────────────────────┐" -ForegroundColor Red
    Write-Host "  │  Missing Required Module: ImportExcel                    │" -ForegroundColor Red
    Write-Host "  │                                                          │" -ForegroundColor Red
    Write-Host "  │  This module is needed to read the Excel spreadsheet.    │" -ForegroundColor Red
    Write-Host "  │                                                          │" -ForegroundColor Red
    Write-Host "  │  To install it, run this command:                        │" -ForegroundColor Red
    Write-Host "  │  Install-Module ImportExcel -Scope CurrentUser -Force    │" -ForegroundColor Yellow
    Write-Host "  │                                                          │" -ForegroundColor Red
    Write-Host "  │  Then come back and try again.                           │" -ForegroundColor Red
    Write-Host "  └──────────────────────────────────────────────────────────┘" -ForegroundColor Red
    Write-Host ""
    return
}
Import-Module ImportExcel

# ─── Load Config ───────────────────────────────────────────────────────────────

$config = Read-MigrationConfig -ConfigPath $ConfigPath
$logFile = Initialize-MigrationLog -LogDirectory $config.logDirectory -ScriptName 'ArchiveRepos'

# ─── Step 1: Locate the Excel File ────────────────────────────────────────────

if ($Interactive) {
    Show-MenuHeader -Title "Archive Repositories"

    Write-Host "  This tool will:" -ForegroundColor White
    Write-Host ""
    Write-Host "    1. Read the Dalptfs01 roster spreadsheet" -ForegroundColor DarkGray
    Write-Host "    2. Show you which repositories will be archived" -ForegroundColor DarkGray
    Write-Host "    3. Ask for your confirmation before creating the archive" -ForegroundColor DarkGray
    Write-Host "    4. Package everything into a zip file with full records" -ForegroundColor DarkGray
    Write-Host ""
    Write-Host "  No changes are made until you explicitly confirm." -ForegroundColor Green
    Write-Host ""
    Write-Host "  ─────────────────────────────────────────────────────" -ForegroundColor DarkGray
    Write-Host "  Step 1 of 3: Locate the spreadsheet" -ForegroundColor Cyan
    Write-Host "  ─────────────────────────────────────────────────────" -ForegroundColor DarkGray
    Write-Host ""
}

if (-not $ExcelPath) {
    $defaultExcel = Join-Path $PSScriptRoot 'excel-docs/Dalptfs01-Collections-MikeFelder.xlsx'
    if (Test-Path $defaultExcel) {
        if ($Interactive) {
            Write-Host "  ✓ Found the spreadsheet automatically:" -ForegroundColor Green
            Write-Host "    $defaultExcel" -ForegroundColor White
            Write-Host ""
        }
        $ExcelPath = $defaultExcel
    }
    elseif ($Interactive) {
        Write-Host "  The spreadsheet was not found in the expected location." -ForegroundColor Yellow
        Write-Host "  Please enter the full path to Dalptfs01-Collections-MikeFelder.xlsx:" -ForegroundColor White
        Write-Host ""
        Write-Host "  Path: " -ForegroundColor Yellow -NoNewline
        $ExcelPath = Read-Host
    }
}

if (-not $ExcelPath -or -not (Test-Path $ExcelPath)) {
    Write-Host ""
    Write-Host "  Could not find the spreadsheet file." -ForegroundColor Red
    Write-Host "  Make sure 'Dalptfs01-Collections-MikeFelder.xlsx' is in the excel-docs folder." -ForegroundColor Yellow
    Write-Host ""
    return
}

# ─── Step 2: Read & Analyze the Spreadsheet ───────────────────────────────────

if ($Interactive) {
    Write-Host "  ─────────────────────────────────────────────────────" -ForegroundColor DarkGray
    Write-Host "  Step 2 of 3: Reading the spreadsheet..." -ForegroundColor Cyan
    Write-Host "  ─────────────────────────────────────────────────────" -ForegroundColor DarkGray
    Write-Host ""
}

Write-MigrationLog -Message "Reading Excel file: $ExcelPath (sheet: $WorksheetName)" -LogFile $logFile -Level INFO

try {
    $excelData = Import-Excel -Path $ExcelPath -WorksheetName $WorksheetName
}
catch {
    Write-Host "  Could not read the spreadsheet." -ForegroundColor Red
    Write-Host "  Error: $($_.Exception.Message)" -ForegroundColor Red
    Write-Host ""
    Write-Host "  Make sure the file is a valid Excel file and is not open in another program." -ForegroundColor Yellow
    Write-Host ""
    return
}

if (-not $excelData -or $excelData.Count -eq 0) {
    Write-Host "  The spreadsheet appears to be empty." -ForegroundColor Red
    Write-Host "  Expected data in worksheet: '$WorksheetName'" -ForegroundColor Yellow
    Write-Host ""
    return
}

Write-MigrationLog -Message "Loaded $($excelData.Count) rows from Excel" -LogFile $logFile -Level INFO

# ─── Classify Rows ─────────────────────────────────────────────────────────────

$archiveRows = [System.Collections.ArrayList]::new()
$nonArchiveRows = [System.Collections.ArrayList]::new()
$dataWarnings = [System.Collections.ArrayList]::new()

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

        # Validate
        if (-not $collection) {
            [void]$dataWarnings.Add("A row is marked for archive but has no Collection name — it will be skipped")
            $entry.Action = 'Skip'
            $entry.Status = 'Skipped'
            $entry.Error = 'Missing collection name in spreadsheet'
            [void]$nonArchiveRows.Add($entry)
            continue
        }
        if (-not $config.collections[$collection]) {
            [void]$dataWarnings.Add("Collection '$collection' is not in your config file — $projectName will fail")
        }

        [void]$archiveRows.Add($entry)
    }
    else {
        $entry.Action = 'NoAction'
        $entry.Status = 'Skipped'
        [void]$nonArchiveRows.Add($entry)
    }
}

# ─── Show the Preview ─────────────────────────────────────────────────────────

$emptyCount = @($archiveRows | Where-Object { $_.IsEmpty -eq 'Yes' }).Count
$withContentCount = $archiveRows.Count - $emptyCount

Write-Host ""
Write-Host "  ╔═══════════════════════════════════════════════════════╗" -ForegroundColor Cyan
Write-Host "  ║              Archive Preview                          ║" -ForegroundColor Cyan
Write-Host "  ╚═══════════════════════════════════════════════════════╝" -ForegroundColor Cyan
Write-Host ""
Write-Host "  The spreadsheet has $($excelData.Count) rows. Here is the breakdown:" -ForegroundColor White
Write-Host ""
Write-Host "    ● $($archiveRows.Count) repositories will be archived into a zip file" -ForegroundColor Yellow
Write-Host "        ├─ $withContentCount with content (will be inventoried)" -ForegroundColor DarkGray
Write-Host "        └─ $emptyCount marked as empty" -ForegroundColor DarkGray
Write-Host "    ● $($nonArchiveRows.Count) repositories are NOT marked for archive (untouched)" -ForegroundColor DarkGray
Write-Host ""

# Show data warnings
if ($dataWarnings.Count -gt 0) {
    Write-Host "  ┌─ Warnings Found ──────────────────────────────────────┐" -ForegroundColor Yellow
    foreach ($w in $dataWarnings | Select-Object -First 10) {
        Write-Host "  │  ⚠ $w" -ForegroundColor Yellow
    }
    if ($dataWarnings.Count -gt 10) {
        Write-Host "  │  ... and $($dataWarnings.Count - 10) more warnings" -ForegroundColor Yellow
    }
    Write-Host "  └───────────────────────────────────────────────────────┘" -ForegroundColor Yellow
    Write-Host ""
}

# Group by collection for display
$archiveByCollection = $archiveRows | Group-Object -Property Collection

Write-Host "  ── Repositories to archive (by collection) ──────────" -ForegroundColor DarkGray
Write-Host ""

foreach ($grp in $archiveByCollection) {
    Write-Host "  Collection: $($grp.Name) ($($grp.Count) repos)" -ForegroundColor White
    foreach ($item in ($grp.Group | Select-Object -First 10)) {
        $label = if ($item.ProjectName) { $item.ProjectName } else { '(collection-level)' }
        $emptyTag = if ($item.IsEmpty -eq 'Yes') { ' [empty]' } else { '' }
        Write-Host "    ○ $label$emptyTag" -ForegroundColor DarkGray
    }
    if ($grp.Count -gt 10) {
        Write-Host "    ... and $($grp.Count - 10) more" -ForegroundColor DarkGray
    }
    Write-Host ""
}

# Always save the preview manifest
$allPreviewActions = [System.Collections.ArrayList]::new()
foreach ($r in $archiveRows) { [void]$allPreviewActions.Add($r) }
foreach ($r in $nonArchiveRows) { [void]$allPreviewActions.Add($r) }

$previewManifestPath = Join-Path $config.outputDirectory "archive-PREVIEW-$(Get-Date -Format 'yyyyMMdd-HHmmss').csv"
$allPreviewActions | Export-Csv -Path $previewManifestPath -NoTypeInformation -Encoding utf8
Write-MigrationLog -Message "Preview manifest written to: $previewManifestPath" -LogFile $logFile -Level SUCCESS

Write-Host "  A detailed preview has been saved to:" -ForegroundColor DarkGray
Write-Host "  $previewManifestPath" -ForegroundColor White
Write-Host "  (You can open this CSV in Excel to review all $($excelData.Count) rows)" -ForegroundColor DarkGray
Write-Host ""

# ─── Step 3: Confirm and Execute ──────────────────────────────────────────────

if ($Interactive) {
    Write-Host "  ─────────────────────────────────────────────────────" -ForegroundColor DarkGray
    Write-Host "  Step 3 of 3: Confirm and run" -ForegroundColor Cyan
    Write-Host "  ─────────────────────────────────────────────────────" -ForegroundColor DarkGray
    Write-Host ""

    if ($archiveRows.Count -eq 0) {
        Write-Host "  There are no repositories to archive." -ForegroundColor Yellow
        Write-Host "  No rows in the spreadsheet have Archive = 'Yes'." -ForegroundColor DarkGray
        Write-Host ""
        return
    }

    Write-Host "  ┌──────────────────────────────────────────────────────────┐" -ForegroundColor Yellow
    Write-Host "  │                                                          │" -ForegroundColor Yellow
    Write-Host "  │  Ready to archive $($archiveRows.Count.ToString().PadRight(4)) repositories.                   │" -ForegroundColor Yellow
    Write-Host "  │                                                          │" -ForegroundColor Yellow
    Write-Host "  │  This will:                                              │" -ForegroundColor Yellow
    Write-Host "  │    • Inventory each project's contents on the server     │" -ForegroundColor Yellow
    Write-Host "  │    • Save metadata for every project                     │" -ForegroundColor Yellow
    Write-Host "  │    • Package everything into one zip file                │" -ForegroundColor Yellow
    Write-Host "  │                                                          │" -ForegroundColor Yellow
    Write-Host "  │  Nothing on the server is deleted or modified.           │" -ForegroundColor Green
    Write-Host "  │  This is a read-only operation — it just creates         │" -ForegroundColor Green
    Write-Host "  │  a local archive copy.                                   │" -ForegroundColor Green
    Write-Host "  │                                                          │" -ForegroundColor Yellow
    Write-Host "  └──────────────────────────────────────────────────────────┘" -ForegroundColor Yellow
    Write-Host ""
    Write-Host "  Do you want to proceed? Type 'yes' to confirm: " -ForegroundColor Yellow -NoNewline
    $confirm = Read-Host

    if ($confirm.Trim().ToLower() -ne 'yes') {
        Write-Host ""
        Write-Host "  Archive cancelled. No changes were made." -ForegroundColor Yellow
        Write-Host "  The preview file is still available at: $previewManifestPath" -ForegroundColor DarkGray
        Write-Host ""
        return
    }

    Write-Host ""
}
elseif ($archiveRows.Count -eq 0) {
    Write-Host "  No repositories to archive. No rows have Archive = 'Yes'." -ForegroundColor Yellow
    return
}

# ─── Execute Archive ──────────────────────────────────────────────────────────

Write-Host "  ╔═══════════════════════════════════════════════════════╗" -ForegroundColor Cyan
Write-Host "  ║              Archiving In Progress                    ║" -ForegroundColor Cyan
Write-Host "  ╚═══════════════════════════════════════════════════════╝" -ForegroundColor Cyan
Write-Host ""

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

    $label = if ($projectName) { "$collection / $projectName" } else { "$collection (collection-level)" }
    Write-Host ""
    Write-Host "  Working on $processedCount of $($archiveRows.Count)..." -ForegroundColor Cyan
    Write-Host "    Archiving: $label" -ForegroundColor White

    try {
        $pat = $config.collections[$collection]?.pat
        if (-not $pat) {
            $friendlyMsg = "The collection '$collection' is not in your configuration file. " +
                "Add it using the Setup Wizard (option 1 on the main menu)."
            throw $friendlyMsg
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

                Write-Host "    Inventoried $($items.Count) items" -ForegroundColor DarkGray
                Write-MigrationLog -Message "Inventoried $($items.Count) items in $label" -LogFile $logFile
            }
            catch {
                # Could not enumerate — write a placeholder
                $errorPath = Join-Path $itemDir 'archive-note.txt'
                "Could not enumerate TFVC contents: $($_.Exception.Message)" | Out-File $errorPath -Encoding utf8
                Write-Host "    Could not list contents (saved a note instead)" -ForegroundColor DarkGray
                Write-MigrationLog -Message "Could not enumerate $label — $($_.Exception.Message)" -LogFile $logFile -Level WARN
            }
        }
        else {
            # Empty project — write placeholder
            $emptyPath = Join-Path $itemDir 'archive-note.txt'
            "Project marked as empty or collection-level entry. No TFVC content to archive." | Out-File $emptyPath -Encoding utf8
            Write-Host "    Marked as empty — saved metadata only" -ForegroundColor DarkGray
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
        Write-Host "    ✓ Done" -ForegroundColor Green
        Write-MigrationLog -Message "Archived: $label" -LogFile $logFile -Level SUCCESS
    }
    catch {
        $item.Status = 'Failed'
        $item.Error = $_.Exception.Message
        $failCount++
        Write-Host "    ✗ Failed: $($_.Exception.Message)" -ForegroundColor Red
        Write-MigrationLog -Message "Failed to archive $label — $($_.Exception.Message)" -LogFile $logFile -Level ERROR
    }
}

# ─── Create Zip ────────────────────────────────────────────────────────────────

Write-Host ""
Write-Host "  Packaging into zip file..." -ForegroundColor Cyan

try {
    Compress-Archive -Path "$archiveDir/*" -DestinationPath $zipPath -Force
    Write-MigrationLog -Message "Zip archive created: $zipPath" -LogFile $logFile -Level SUCCESS
    Write-Host "  ✓ Zip created successfully" -ForegroundColor Green

    # Clean up staging directory
    Remove-Item -Path $archiveDir -Recurse -Force
}
catch {
    Write-MigrationLog -Message "Failed to create zip: $($_.Exception.Message)" -LogFile $logFile -Level ERROR
    Write-Host "  ✗ Could not create zip file: $($_.Exception.Message)" -ForegroundColor Red
    Write-Host "    The unzipped files are still available at: $archiveDir" -ForegroundColor Yellow
}

# ─── Final Report ─────────────────────────────────────────────────────────────

$totalDuration = (Get-Date) - $startTime

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

# Write JSON report
$reportPath = Join-Path $config.outputDirectory "archive-report-$(Get-Date -Format 'yyyyMMdd-HHmmss').json"
$report = @{
    startTime       = $startTime.ToString('o')
    endTime         = (Get-Date).ToString('o')
    duration        = $totalDuration.ToString('hh\:mm\:ss')
    excelFile       = $ExcelPath
    worksheet       = $WorksheetName
    zipFile         = $zipPath
    totalRows       = $excelData.Count
    archiveCount    = $archiveRows.Count
    successCount    = $successCount
    failCount       = $failCount
    nonArchiveCount = $nonArchiveRows.Count
}
$report | ConvertTo-Json -Depth 5 | Out-File $reportPath -Encoding utf8

Write-MigrationLog -Message "Archive manifest: $manifestPath" -LogFile $logFile -Level SUCCESS
Write-MigrationLog -Message "Report: $reportPath" -LogFile $logFile -Level INFO

# Display results
Write-Host ""
Write-Host ""
Write-Host "  ╔═══════════════════════════════════════════════════════╗" -ForegroundColor Cyan
Write-Host "  ║              Archive Complete                         ║" -ForegroundColor Cyan
Write-Host "  ╚═══════════════════════════════════════════════════════╝" -ForegroundColor Cyan
Write-Host ""
Write-Host "  Time elapsed:  $($totalDuration.ToString('hh\:mm\:ss'))" -ForegroundColor White
Write-Host ""
Write-Host "  Results:" -ForegroundColor White
Write-Host "    ✓ Archived:      $successCount" -ForegroundColor Green
if ($failCount -gt 0) {
    Write-Host "    ✗ Failed:        $failCount" -ForegroundColor Red
}
Write-Host "    ○ Not archived:  $($nonArchiveRows.Count) (not marked for archive)" -ForegroundColor DarkGray
Write-Host ""

if (Test-Path $zipPath) {
    $zipSize = (Get-Item $zipPath).Length
    $sizeLabel = if ($zipSize -gt 1MB) { "$([math]::Round($zipSize / 1MB, 1)) MB" } else { "$([math]::Round($zipSize / 1KB, 0)) KB" }
    Write-Host "  Zip file:  $zipPath ($sizeLabel)" -ForegroundColor White
    Write-Host ""
}

if ($failCount -gt 0) {
    Write-Host "  ┌─ Items that failed ───────────────────────────────────┐" -ForegroundColor Red
    foreach ($f in ($archiveRows | Where-Object { $_.Status -eq 'Failed' }) | Select-Object -First 15) {
        $label = if ($f.ProjectName) { "$($f.Collection) / $($f.ProjectName)" } else { $f.Collection }
        Write-Host "  │  ✗ $label" -ForegroundColor Red
        Write-Host "  │    $($f.Error)" -ForegroundColor DarkGray
    }
    if ($failCount -gt 15) {
        Write-Host "  │  ... and $($failCount - 15) more (see manifest for full list)" -ForegroundColor Red
    }
    Write-Host "  └───────────────────────────────────────────────────────┘" -ForegroundColor Red
    Write-Host ""
}

Write-Host "  Output files:" -ForegroundColor White
Write-Host "    Manifest (CSV):  $manifestPath" -ForegroundColor DarkGray
Write-Host "    Report (JSON):   $reportPath" -ForegroundColor DarkGray
Write-Host "    Log file:        $logFile" -ForegroundColor DarkGray
Write-Host ""
Write-Host "  The manifest CSV can be opened in Excel for a complete" -ForegroundColor DarkGray
Write-Host "  record of every repository and what happened to it." -ForegroundColor DarkGray
Write-Host ""
Write-Host ""
