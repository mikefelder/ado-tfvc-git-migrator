<#
.SYNOPSIS
  Pull latest GitHub Copilot enterprise usage report, aggregate by user, optionally map to teams/cost centers,
  and export summarized CSVs.

.NOTES
  - Built for the newer Copilot metrics/reporting model that returns download_links to NDJSON report files.
  - You may need to adjust field names after inspecting your tenant's actual NDJSON schema.
  - Cost estimation in this script is an optional approximation that YOU control with the rate inputs below.
#>#
# Example to run:
# .\Get-GHCopilotUsage.ps1 `
#  -Enterprise "your-enterprise" `
#  -GitHubToken "ghp_xxx"


param(
    [Parameter(Mandatory = $true)]
    [string]$Enterprise,

    [Parameter(Mandatory = $true)]
    [string]$GitHubToken,

    [string]$OutputFolder = ".\ghcp-usage-output",

    # Optional mapping CSV with columns like:
    # user_login,display_name,team,cost_center,manager
    [string]$UserMapCsv = "",

    # Optional: estimated cost inputs (your own approximation)
    [decimal]$CostPerToken = 0.0,         # e.g. 0.000001
    [decimal]$CostPerRequest = 0.0,       # e.g. 0.01 if you want a crude request-based estimate
    [switch]$EstimateCost
)

# ----------------------------
# Helper functions
# ----------------------------

function Write-Section {
    param([string]$Message)
    Write-Host ""
    Write-Host "==== $Message ====" -ForegroundColor Cyan
}

function Get-JsonValue {
    param(
        [Parameter(Mandatory=$true)]$Object,
        [Parameter(Mandatory=$true)][string[]]$Names
    )
    foreach ($name in $Names) {
        if ($Object.PSObject.Properties.Name -contains $name) {
            return $Object.$name
        }
    }
    return $null
}

function To-DecimalSafe {
    param($Value)
    if ($null -eq $Value -or $Value -eq "") { return [decimal]0 }
    try { return [decimal]$Value } catch { return [decimal]0 }
}

function To-StringSafe {
    param($Value)
    if ($null -eq $Value) { return "" }
    return [string]$Value
}

# ----------------------------
# Setup
# ----------------------------

$ErrorActionPreference = "Stop"

if (-not (Test-Path $OutputFolder)) {
    New-Item -ItemType Directory -Path $OutputFolder | Out-Null
}

$headers = @{
    Authorization = "Bearer $GitHubToken"
    Accept        = "application/vnd.github+json"
    "User-Agent"  = "ghcp-usage-report-script"
}

Write-Section "Requesting latest enterprise 28-day report metadata"

# Current report model uses NDJSON download links for the latest 28-day report
$reportUrl = "https://api.github.com/enterprises/$Enterprise/copilot/metrics/reports/enterprise-28-day/latest"
$reportMeta = Invoke-RestMethod -Uri $reportUrl -Headers $headers -Method Get

if (-not $reportMeta.download_links -or $reportMeta.download_links.Count -eq 0) {
    throw "No download_links were returned. Check token access, enterprise name, and whether metrics/reporting are enabled."
}

$reportMeta | ConvertTo-Json -Depth 10 | Out-File (Join-Path $OutputFolder "report-metadata.json") -Encoding utf8

# ----------------------------
# Download + parse NDJSON
# ----------------------------

Write-Section "Downloading and parsing NDJSON report files"

$rawRows = New-Object System.Collections.Generic.List[object]
$fileIndex = 0

foreach ($link in $reportMeta.download_links) {
    $fileIndex++
    Write-Host "Downloading file $fileIndex of $($reportMeta.download_links.Count)..."

    $resp = Invoke-WebRequest -Uri $link -Headers $headers -Method Get
    $rawText = $resp.Content

    $rawFile = Join-Path $OutputFolder ("raw-report-" + $fileIndex + ".ndjson")
    $rawText | Out-File $rawFile -Encoding utf8

    $lines = $rawText -split "`r?`n"
    foreach ($line in $lines) {
        if ([string]::IsNullOrWhiteSpace($line)) { continue }
        try {
            $obj = $line | ConvertFrom-Json
            $rawRows.Add($obj)
        }
        catch {
            Write-Warning "Skipping malformed NDJSON line in file $fileIndex"
        }
    }
}

if ($rawRows.Count -eq 0) {
    throw "No rows were parsed from the NDJSON report files."
}

Write-Host "Parsed $($rawRows.Count) rows." -ForegroundColor Green

# ----------------------------
# Normalize rows
# ----------------------------

Write-Section "Normalizing schema"

$normalized = foreach ($row in $rawRows) {

    # Common possible user identifiers
    $userLogin   = To-StringSafe (Get-JsonValue -Object $row -Names @("user_login","user","login","actor","username"))
    $displayName = To-StringSafe (Get-JsonValue -Object $row -Names @("display_name","user_name","name"))
    $orgName     = To-StringSafe (Get-JsonValue -Object $row -Names @("organization","org","organization_login"))
    $dateValue   = To-StringSafe (Get-JsonValue -Object $row -Names @("day","date","usage_date","timestamp"))

    # Common possible metrics
    $totalRequests = To-DecimalSafe (Get-JsonValue -Object $row -Names @("total_requests","requests","request_count"))
    $totalTokens   = To-DecimalSafe (Get-JsonValue -Object $row -Names @("total_tokens","tokens","token_count"))
    $promptTokens  = To-DecimalSafe (Get-JsonValue -Object $row -Names @("prompt_tokens","input_tokens"))
    $outputTokens  = To-DecimalSafe (Get-JsonValue -Object $row -Names @("completion_tokens","output_tokens","generated_tokens"))
    $activeDays    = To-DecimalSafe (Get-JsonValue -Object $row -Names @("active_days","days_active"))

    # If total tokens aren't present but prompt/output are
    if ($totalTokens -eq 0 -and ($promptTokens -gt 0 -or $outputTokens -gt 0)) {
        $totalTokens = $promptTokens + $outputTokens
    }

    [PSCustomObject]@{
        user_login      = $userLogin
        display_name    = $displayName
        organization    = $orgName
        date_value      = $dateValue
        requests        = $totalRequests
        total_tokens    = $totalTokens
        prompt_tokens   = $promptTokens
        output_tokens   = $outputTokens
        active_days     = $activeDays
        raw             = $row
    }
}

$normalizedCsv = Join-Path $OutputFolder "normalized-rows.csv"
$normalized | Select-Object user_login,display_name,organization,date_value,requests,total_tokens,prompt_tokens,output_tokens,active_days |
    Export-Csv -NoTypeInformation -Path $normalizedCsv -Encoding utf8

# ----------------------------
# Aggregate by user
# ----------------------------

Write-Section "Aggregating by user"

$userSummary = $normalized |
    Where-Object { -not [string]::IsNullOrWhiteSpace($_.user_login) } |
    Group-Object user_login |
    ForEach-Object {
        $group = $_.Group
        $sample = $group | Select-Object -First 1

        $sumRequests    = ($group | Measure-Object -Property requests -Sum).Sum
        $sumTokens      = ($group | Measure-Object -Property total_tokens -Sum).Sum
        $sumPrompt      = ($group | Measure-Object -Property prompt_tokens -Sum).Sum
        $sumOutput      = ($group | Measure-Object -Property output_tokens -Sum).Sum
        $sumActiveDays  = ($group | Measure-Object -Property active_days -Sum).Sum

        $estCost = [decimal]0
        if ($EstimateCost) {
            $estCost = ([decimal]$sumTokens * $CostPerToken) + ([decimal]$sumRequests * $CostPerRequest)
        }

        [PSCustomObject]@{
            user_login        = $_.Name
            display_name      = $sample.display_name
            organization      = $sample.organization
            total_requests    = [decimal]$sumRequests
            total_tokens      = [decimal]$sumTokens
            prompt_tokens     = [decimal]$sumPrompt
            output_tokens     = [decimal]$sumOutput
            active_days       = [decimal]$sumActiveDays
            estimated_cost    = [decimal]$estCost
        }
    } |
    Sort-Object total_tokens -Descending

# ----------------------------
# Join optional mapping CSV
# ----------------------------

if (-not [string]::IsNullOrWhiteSpace($UserMapCsv) -and (Test-Path $UserMapCsv)) {
    Write-Section "Joining user/team/cost center mapping"

    $mapRows = Import-Csv $UserMapCsv
    $mapIndex = @{}
    foreach ($m in $mapRows) {
        if ($m.user_login) {
            $mapIndex[$m.user_login.ToLower()] = $m
        }
    }

    $userSummary = $userSummary | ForEach-Object {
        $key = $_.user_login.ToLower()
        $map = $null
        if ($mapIndex.ContainsKey($key)) { $map = $mapIndex[$key] }

        [PSCustomObject]@{
            user_login      = $_.user_login
            display_name    = if ($map -and $map.display_name) { $map.display_name } else { $_.display_name }
            organization    = $_.organization
            team            = if ($map) { $map.team } else { "" }
            cost_center     = if ($map) { $map.cost_center } else { "" }
            manager         = if ($map) { $map.manager } else { "" }
            total_requests  = $_.total_requests
            total_tokens    = $_.total_tokens
            prompt_tokens   = $_.prompt_tokens
            output_tokens   = $_.output_tokens
            active_days     = $_.active_days
            estimated_cost  = $_.estimated_cost
        }
    }
}
else {
    $userSummary = $userSummary | ForEach-Object {
        [PSCustomObject]@{
            user_login      = $_.user_login
            display_name    = $_.display_name
            organization    = $_.organization
            team            = ""
            cost_center     = ""
            manager         = ""
            total_requests  = $_.total_requests
            total_tokens    = $_.total_tokens
            prompt_tokens   = $_.prompt_tokens
            output_tokens   = $_.output_tokens
            active_days     = $_.active_days
            estimated_cost  = $_.estimated_cost
        }
    }
}

# ----------------------------
# Aggregate by team
# ----------------------------

Write-Section "Aggregating by team"

$teamSummary = $userSummary |
    Group-Object team |
    ForEach-Object {
        $group = $_.Group
        [PSCustomObject]@{
            team            = if ([string]::IsNullOrWhiteSpace($_.Name)) { "(unmapped)" } else { $_.Name }
            users           = ($group | Measure-Object).Count
            total_requests  = ($group | Measure-Object -Property total_requests -Sum).Sum
            total_tokens    = ($group | Measure-Object -Property total_tokens -Sum).Sum
            prompt_tokens   = ($group | Measure-Object -Property prompt_tokens -Sum).Sum
            output_tokens   = ($group | Measure-Object -Property output_tokens -Sum).Sum
            estimated_cost  = ($group | Measure-Object -Property estimated_cost -Sum).Sum
        }
    } |
    Sort-Object total_tokens -Descending

# ----------------------------
# Export
# ----------------------------

Write-Section "Writing CSV outputs"

$userCsv = Join-Path $OutputFolder "copilot-user-usage-summary.csv"
$teamCsv = Join-Path $OutputFolder "copilot-team-usage-summary.csv"

$userSummary | Export-Csv -NoTypeInformation -Path $userCsv -Encoding utf8
$teamSummary | Export-Csv -NoTypeInformation -Path $teamCsv -Encoding utf8

# ----------------------------
# Console output
# ----------------------------

Write-Section "Top 25 users by total tokens"
$userSummary |
    Sort-Object total_tokens -Descending |
    Select-Object -First 25 user_login,display_name,team,total_requests,total_tokens,estimated_cost |
    Format-Table -AutoSize

Write-Section "Top teams by total tokens"
$teamSummary |
    Select-Object -First 15 team,users,total_requests,total_tokens,estimated_cost |
    Format-Table -AutoSize

Write-Section "Done"
Write-Host "User summary: $userCsv"
Write-Host "Team summary: $teamCsv"
Write-Host "Metadata saved in: $OutputFolder"
``