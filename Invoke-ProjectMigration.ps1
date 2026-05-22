#Requires -Version 7.0
<#
.SYNOPSIS
    Interactive per-project mirror picker. Lists every project in the SOURCE
    collection alongside its migration status against a TARGET collection
    (defaults to MDR-GAMS-ADO when present in the config), lets the user pick
    a single project via arrow keys or by typing the row number, and then
    invokes Mirror-AdoCollection.ps1 for just that one project.

.DESCRIPTION
    Workflow:
      1. Pick source + target collections (auto-selected if unambiguous).
      2. Query both endpoints in parallel-ish: list source projects + repos,
         list target projects + repos.
      3. Compute a per-project status:
           • Not migrated  — project missing in target
           • Migrated      — all source repos exist in target (and target
                             repo count >= source repo count)
           • Partial       — project exists in target but some repos are
                             missing or extra
           • Empty         — source project has no Git repos
      4. Render an interactive table and accept either:
            ↑ / ↓ + Enter           — keyboard selection
            <digits> + Enter        — jump straight to a row
            Q / Esc                 — cancel
            R                       — refresh statuses
            A                       — toggle "show all" vs "hide migrated"
      5. Hand off to Mirror-AdoCollection.ps1 with -SourceProject <picked>.

.PARAMETER ConfigPath
    Path to migration-config.json.

.PARAMETER SourceCollection
    Optional. If omitted and only one source-only collection is defined, it is
    auto-picked. Otherwise the user is prompted with Select-AdoCollection.

.PARAMETER TargetCollection
    Optional. Defaults to 'MDR-GAMS-ADO' when present in the config; otherwise
    the user is prompted.

.PARAMETER DryRun
    Forwarded to Mirror-AdoCollection.ps1 for the picked project.

.PARAMETER NoProgress
    Forwarded to Mirror-AdoCollection.ps1.

.PARAMETER Force
    Forwarded to Mirror-AdoCollection.ps1 (skips the per-mirror "type yes"
    prompt). The picker's own final confirmation is still shown unless this
    switch is set.

.EXAMPLE
    ./Invoke-ProjectMigration.ps1 -ConfigPath ./config/migration-config.json

.EXAMPLE
    ./Invoke-ProjectMigration.ps1 -ConfigPath ./config/migration-config.json `
        -SourceCollection GAMS-GIT-Repos -TargetCollection MDR-GAMS-ADO -DryRun
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory)]
    [string]$ConfigPath,

    [string]$SourceCollection,
    [string]$TargetCollection,

    [switch]$DryRun,
    [switch]$NoProgress,
    [switch]$Force
)

$ErrorActionPreference = 'Stop'
Import-Module "$PSScriptRoot/modules/AdoTfvcMigrator.psm1" -Force

# ─── Load Config ───────────────────────────────────────────────────────────────

$config = Read-MigrationConfig -ConfigPath $ConfigPath

# ─── Pick Source + Target Collections ────────────────────────────────────────

$collectionNames = @($config.collections.Keys | Sort-Object)
if ($collectionNames.Count -lt 2) {
    throw "Migrate-Project mode needs at least two collections in the config (a source and a target)."
}

# Auto-pick target if MDR-GAMS-ADO is present and no override provided
if (-not $TargetCollection) {
    if ($collectionNames -contains 'MDR-GAMS-ADO') {
        $TargetCollection = 'MDR-GAMS-ADO'
    }
    else {
        Show-MenuHeader -Title 'Select TARGET collection (Azure DevOps Services org)'
        $TargetCollection = Select-AdoCollection -Config $config -Prompt 'Select TARGET collection'
        if (-not $TargetCollection) { Write-Host 'Cancelled.' -ForegroundColor Yellow; return }
    }
}

if (-not $SourceCollection) {
    $sourceCandidates = @($collectionNames | Where-Object { $_ -ne $TargetCollection })
    if ($sourceCandidates.Count -eq 1) {
        $SourceCollection = $sourceCandidates[0]
    }
    else {
        Show-MenuHeader -Title 'Select SOURCE collection'
        $SourceCollection = Select-AdoCollection -Config $config -Prompt 'Select SOURCE collection'
        if (-not $SourceCollection) { Write-Host 'Cancelled.' -ForegroundColor Yellow; return }
    }
}

if ($SourceCollection -eq $TargetCollection) {
    throw "Source and target collections must differ ('$SourceCollection')."
}

# ─── Resolve Endpoints + PATs ────────────────────────────────────────────────

$sourceUrl = Get-CollectionServerUrl -Config $config -Collection $SourceCollection
$targetUrl = Get-CollectionServerUrl -Config $config -Collection $TargetCollection
$sourcePat = Get-CollectionPat        -Config $config -Collection $SourceCollection
$targetPat = Get-CollectionPat        -Config $config -Collection $TargetCollection

# ─── Validate Both Endpoints ─────────────────────────────────────────────────

Write-Host ""
Write-Host "Validating SOURCE connection ($SourceCollection)..." -ForegroundColor Cyan
$srcCheck = Test-AdoConnection -ServerUrl $sourceUrl -Collection $SourceCollection -Pat $sourcePat
if (-not $srcCheck.Connected) {
    throw "Cannot reach source '$SourceCollection': $($srcCheck.Error)"
}
Write-Host "  ✓ Source OK ($($srcCheck.ProjectCount) projects)" -ForegroundColor Green

Write-Host "Validating TARGET connection ($TargetCollection)..." -ForegroundColor Cyan
$tgtCheck = Test-AdoConnection -ServerUrl $targetUrl -Collection $TargetCollection -Pat $targetPat
if (-not $tgtCheck.Connected) {
    throw "Cannot reach target '$TargetCollection': $($tgtCheck.Error)"
}
Write-Host "  ✓ Target OK ($($tgtCheck.ProjectCount) existing projects)" -ForegroundColor Green

# ─── Build Status Table ──────────────────────────────────────────────────────

function Get-ProjectStatusRows {
    [CmdletBinding()]
    param()

    Write-Host ""
    Write-Host "Enumerating source projects + repos..." -ForegroundColor Cyan
    $srcProjects = Get-AdoProjects -ServerUrl $sourceUrl -Collection $SourceCollection -Pat $sourcePat |
        Sort-Object name

    Write-Host "Enumerating target projects..." -ForegroundColor Cyan
    $tgtProjects = Get-AdoProjects -ServerUrl $targetUrl -Collection $TargetCollection -Pat $targetPat
    $tgtByName = @{}
    foreach ($p in $tgtProjects) { $tgtByName[$p.name] = $p }

    $rows = [System.Collections.Generic.List[object]]::new()
    $total = $srcProjects.Count
    $i = 0
    foreach ($p in $srcProjects) {
        $i++
        if (-not $NoProgress) {
            Write-Progress -Id 50 -Activity 'Scanning source/target repos' `
                -Status ("[{0}/{1}] {2}" -f $i, $total, $p.name) `
                -PercentComplete ([int](($i * 100) / [Math]::Max(1, $total)))
        }

        # Source repos
        try {
            $srcRepos = @(Get-AdoGitRepositories -ServerUrl $sourceUrl -Collection $SourceCollection `
                    -ProjectIdOrName $p.id -Pat $sourcePat)
        }
        catch {
            $srcRepos = @()
        }

        # Target project + repos
        $tgtProj = $tgtByName[$p.name]
        $tgtRepos = @()
        if ($tgtProj) {
            try {
                $tgtRepos = @(Get-AdoGitRepositories -ServerUrl $targetUrl -Collection $TargetCollection `
                        -ProjectIdOrName $tgtProj.id -Pat $targetPat)
            }
            catch {
                $tgtRepos = @()
            }
        }

        # Status decision
        $srcCount = $srcRepos.Count
        $tgtCount = $tgtRepos.Count
        $status = if (-not $tgtProj) {
            if ($srcCount -eq 0) { 'Empty (not migrated)' } else { 'Not migrated' }
        }
        elseif ($srcCount -eq 0) {
            'Empty (target exists)'
        }
        else {
            # Compare by repo name (case-insensitive)
            $srcNames = @($srcRepos | ForEach-Object { $_.name.ToLowerInvariant() })
            $tgtNames = @($tgtRepos | ForEach-Object { $_.name.ToLowerInvariant() })
            $missing  = @($srcNames | Where-Object { $_ -notin $tgtNames })
            if ($missing.Count -eq 0) { 'Migrated' } else { "Partial ($($missing.Count) missing)" }
        }

        $rows.Add([PSCustomObject]@{
                Name        = $p.name
                SrcRepos    = $srcCount
                TgtRepos    = $tgtCount
                Status      = $status
                ProjectId   = $p.id
                TargetExists = [bool]$tgtProj
            })
    }
    if (-not $NoProgress) { Write-Progress -Id 50 -Activity ' ' -Completed }
    return , $rows.ToArray()
}

$rows = Get-ProjectStatusRows

if (-not $rows -or $rows.Count -eq 0) {
    Write-Host ""
    Write-Host "No projects found in source collection '$SourceCollection'." -ForegroundColor Yellow
    return
}

# ─── Interactive Selector ────────────────────────────────────────────────────

function Get-StatusColor {
    param([string]$Status)
    switch -Regex ($Status) {
        '^Migrated$'             { 'Green';      break }
        '^Partial'               { 'Yellow';     break }
        '^Not migrated$'         { 'Red';        break }
        '^Empty'                 { 'DarkGray';   break }
        default                  { 'White' }
    }
}

function Get-StatusSymbol {
    param([string]$Status)
    switch -Regex ($Status) {
        '^Migrated$'             { '[OK]'    ; break }
        '^Partial'               { '[~]'     ; break }
        '^Not migrated$'         { '[ ]'     ; break }
        '^Empty'                 { '[--]'    ; break }
        default                  { '[?]' }
    }
}

function Show-ProjectPicker {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)] [object[]]$AllRows,
        [string]$SourceLabel,
        [string]$TargetLabel
    )

    $hideMigrated = $false
    $digitBuffer  = ''
    $selectedIdx  = 0
    $topIdx       = 0

    # Find the first not-migrated row to pre-select for convenience
    for ($i = 0; $i -lt $AllRows.Count; $i++) {
        if ($AllRows[$i].Status -notmatch '^(Migrated|Empty)') { $selectedIdx = $i; break }
    }

    while ($true) {
        $visibleRows = if ($hideMigrated) {
            @($AllRows | Where-Object { $_.Status -notmatch '^Migrated$' })
        }
        else { $AllRows }

        if ($visibleRows.Count -eq 0) {
            [Console]::Clear()
            Write-Host ""
            Write-Host "  No projects match the current filter (all are 'Migrated')." -ForegroundColor Yellow
            Write-Host "  Press [A] to show migrated rows, or [Q] to cancel." -ForegroundColor DarkGray
        }
        else {
            if ($selectedIdx -ge $visibleRows.Count) { $selectedIdx = $visibleRows.Count - 1 }
            if ($selectedIdx -lt 0) { $selectedIdx = 0 }

            # Compute scroll window
            $reservedLines = 12   # header + footer + hint + buffer
            $winHeight = [Math]::Max(5, [Console]::WindowHeight - $reservedLines)
            if ($selectedIdx -lt $topIdx) { $topIdx = $selectedIdx }
            if ($selectedIdx -ge ($topIdx + $winHeight)) { $topIdx = $selectedIdx - $winHeight + 1 }
            if ($topIdx -lt 0) { $topIdx = 0 }
            $endIdx = [Math]::Min($visibleRows.Count - 1, $topIdx + $winHeight - 1)

            # Render
            [Console]::Clear()
            Write-Host ""
            Write-Host "═══════════════════════════════════════════════════════════════════════" -ForegroundColor Cyan
            Write-Host "  Migrate ONE Project — pick a source project to mirror to the target" -ForegroundColor Cyan
            Write-Host "═══════════════════════════════════════════════════════════════════════" -ForegroundColor Cyan
            Write-Host ""
            Write-Host ("  Source : {0}" -f $SourceLabel) -ForegroundColor White
            Write-Host ("  Target : {0}" -f $TargetLabel) -ForegroundColor White
            $filterMsg = if ($hideMigrated) { ' (hiding already-migrated)' } else { '' }
            Write-Host ("  Projects: {0}{1}" -f $visibleRows.Count, $filterMsg) -ForegroundColor DarkGray
            Write-Host ""

            # Column widths
            $maxName = ($visibleRows | ForEach-Object { $_.Name.Length } | Measure-Object -Maximum).Maximum
            if ($maxName -lt 20) { $maxName = 20 }
            if ($maxName -gt 50) { $maxName = 50 }
            $numWidth = [Math]::Max(3, ($visibleRows.Count.ToString().Length))

            # Header
            $header = ('   {0}  {1}  {2}  {3}  {4}' -f `
                ('#'.PadLeft($numWidth)),
                'Project'.PadRight($maxName),
                'Src'.PadLeft(4),
                'Tgt'.PadLeft(4),
                'Status')
            Write-Host $header -ForegroundColor DarkCyan
            Write-Host ('   ' + ('─' * ($header.Length - 3))) -ForegroundColor DarkGray

            for ($i = $topIdx; $i -le $endIdx; $i++) {
                $row = $visibleRows[$i]
                $marker = if ($i -eq $selectedIdx) { '▶' } else { ' ' }
                $rowNum = ($i + 1).ToString().PadLeft($numWidth)
                $name = if ($row.Name.Length -gt $maxName) {
                    $row.Name.Substring(0, $maxName - 1) + '…'
                }
                else { $row.Name.PadRight($maxName) }
                $sym = Get-StatusSymbol -Status $row.Status
                $clr = Get-StatusColor  -Status $row.Status

                $line = (' {0} {1}  {2}  {3}  {4}  {5} {6}' -f `
                    $marker, $rowNum, $name,
                    $row.SrcRepos.ToString().PadLeft(4),
                    $row.TgtRepos.ToString().PadLeft(4),
                    $sym, $row.Status)

                if ($i -eq $selectedIdx) {
                    Write-Host $line -ForegroundColor Black -BackgroundColor White
                }
                else {
                    Write-Host $line -ForegroundColor $clr
                }
            }

            # Scroll hint
            if ($visibleRows.Count -gt $winHeight) {
                Write-Host ("   ── showing {0}-{1} of {2} ──" -f ($topIdx + 1), ($endIdx + 1), $visibleRows.Count) -ForegroundColor DarkGray
            }
        }

        Write-Host ""
        Write-Host "  ↑/↓ move   Enter select   1-$($AllRows.Count) jump   [A] toggle filter   [R] refresh   [Q] cancel" -ForegroundColor DarkGray
        if ($digitBuffer) {
            Write-Host ("  Jump to: {0}_" -f $digitBuffer) -ForegroundColor Yellow
        }
        else {
            Write-Host '  ' -NoNewline
        }

        $key = [Console]::ReadKey($true)

        switch ($key.Key) {
            'UpArrow' {
                $digitBuffer = ''
                if ($selectedIdx -gt 0) { $selectedIdx-- }
            }
            'DownArrow' {
                $digitBuffer = ''
                if ($selectedIdx -lt ($visibleRows.Count - 1)) { $selectedIdx++ }
            }
            'PageUp' {
                $digitBuffer = ''
                $selectedIdx = [Math]::Max(0, $selectedIdx - 10)
            }
            'PageDown' {
                $digitBuffer = ''
                $selectedIdx = [Math]::Min($visibleRows.Count - 1, $selectedIdx + 10)
            }
            'Home' {
                $digitBuffer = ''
                $selectedIdx = 0
            }
            'End' {
                $digitBuffer = ''
                $selectedIdx = $visibleRows.Count - 1
            }
            'Backspace' {
                if ($digitBuffer.Length -gt 0) {
                    $digitBuffer = $digitBuffer.Substring(0, $digitBuffer.Length - 1)
                }
            }
            'Escape' { return $null }
            'Enter' {
                if ($digitBuffer) {
                    $jump = [int]$digitBuffer
                    $digitBuffer = ''
                    # Numeric entries refer to ALL rows (1-based), not the filtered view
                    if ($jump -ge 1 -and $jump -le $AllRows.Count) {
                        return $AllRows[$jump - 1]
                    }
                    # Fall through: invalid number, just stay in loop
                }
                else {
                    if ($visibleRows.Count -gt 0) {
                        return $visibleRows[$selectedIdx]
                    }
                }
            }
            default {
                $ch = $key.KeyChar
                if ($ch -match '[0-9]') {
                    # Cap buffer to length needed for max index
                    $maxLen = $AllRows.Count.ToString().Length
                    if ($digitBuffer.Length -lt $maxLen) {
                        $digitBuffer += $ch
                    }
                }
                elseif ($ch -in @('q', 'Q')) {
                    return $null
                }
                elseif ($ch -in @('a', 'A')) {
                    $digitBuffer = ''
                    $hideMigrated = -not $hideMigrated
                    $selectedIdx = 0
                    $topIdx = 0
                }
                elseif ($ch -in @('r', 'R')) {
                    $digitBuffer = ''
                    return '__REFRESH__'
                }
            }
        }
    }
}

# Selector loop — supports refresh
while ($true) {
    $sourceLabel = ('{0} ({1})' -f $SourceCollection, $sourceUrl)
    $targetLabel = ('{0} ({1})' -f $TargetCollection, $targetUrl)
    $picked = Show-ProjectPicker -AllRows $rows -SourceLabel $sourceLabel -TargetLabel $targetLabel

    if ($null -eq $picked) {
        Write-Host ""
        Write-Host "Cancelled." -ForegroundColor Yellow
        return
    }

    if ($picked -is [string] -and $picked -eq '__REFRESH__') {
        $rows = Get-ProjectStatusRows
        continue
    }

    break
}

# ─── Confirm + Hand-Off to Mirror-AdoCollection ──────────────────────────────

[Console]::Clear()
Show-MenuHeader -Title ('Migrate project: {0}' -f $picked.Name)
Write-Host ("  Source  : {0}/{1}/{2}" -f $sourceUrl, $SourceCollection, $picked.Name) -ForegroundColor White
Write-Host ("  Target  : {0}/{1}/{2}" -f $targetUrl, $TargetCollection, $picked.Name) -ForegroundColor White
Write-Host ("  Status  : {0}" -f $picked.Status) -ForegroundColor (Get-StatusColor -Status $picked.Status)
Write-Host ("  Repos   : source={0}  target={1}" -f $picked.SrcRepos, $picked.TgtRepos) -ForegroundColor White
if ($DryRun) {
    Write-Host "  Mode    : DRY RUN (no changes will be made)" -ForegroundColor Yellow
}
Write-Host ""

if (-not $Force) {
    Write-Host "  Proceed with mirror? [y/N]: " -ForegroundColor Yellow -NoNewline
    $ans = Read-Host
    if ($ans.Trim() -notmatch '^[Yy]') {
        Write-Host "Cancelled." -ForegroundColor Yellow
        return
    }
}

# Forward to Mirror-AdoCollection with single-project filter
$mirrorArgs = @{
    ConfigPath       = $ConfigPath
    SourceCollection = $SourceCollection
    TargetCollection = $TargetCollection
    SourceProject    = $picked.Name
    Force            = $true   # we already asked above; let mirror skip its own prompt
}
if ($DryRun)     { $mirrorArgs['DryRun']     = $true }
if ($NoProgress) { $mirrorArgs['NoProgress'] = $true }

& "$PSScriptRoot/Mirror-AdoCollection.ps1" @mirrorArgs
