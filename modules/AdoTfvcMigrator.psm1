#Requires -Version 7.0
<#
.SYNOPSIS
    Shared functions for the ADO TFVC-to-Git migration toolkit.
#>

# ─── Configuration ─────────────────────────────────────────────────────────────

function Read-MigrationConfig {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$ConfigPath
    )

    if (-not (Test-Path $ConfigPath)) {
        throw "Config file not found: $ConfigPath. Copy config/migration-config.example.json and edit it."
    }

    $config = Get-Content $ConfigPath -Raw | ConvertFrom-Json -AsHashtable
    # Resolve relative paths
    $configDir = Split-Path $ConfigPath -Parent
    if ($config.outputDirectory -and -not [IO.Path]::IsPathRooted($config.outputDirectory)) {
        $config.outputDirectory = Join-Path $configDir $config.outputDirectory
    }
    if ($config.logDirectory -and -not [IO.Path]::IsPathRooted($config.logDirectory)) {
        $config.logDirectory = Join-Path $configDir $config.logDirectory
    }
    if ($config.authorMappingFile -and -not [IO.Path]::IsPathRooted($config.authorMappingFile)) {
        $config.authorMappingFile = Join-Path $configDir $config.authorMappingFile
    }

    # Ensure directories exist
    foreach ($dir in @($config.outputDirectory, $config.logDirectory)) {
        if ($dir -and -not (Test-Path $dir)) {
            New-Item -ItemType Directory -Path $dir -Force | Out-Null
        }
    }

    return $config
}

# ─── Logging ───────────────────────────────────────────────────────────────────

function Initialize-MigrationLog {
    [CmdletBinding()]
    param(
        [string]$LogDirectory,
        [string]$ScriptName
    )

    $timestamp = Get-Date -Format 'yyyyMMdd-HHmmss'
    $logFile = Join-Path $LogDirectory "${ScriptName}_${timestamp}.log"

    if (-not (Test-Path $LogDirectory)) {
        New-Item -ItemType Directory -Path $LogDirectory -Force | Out-Null
    }

    return $logFile
}

function Write-MigrationLog {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$Message,
        [string]$LogFile,
        [ValidateSet('INFO', 'WARN', 'ERROR', 'SUCCESS')]
        [string]$Level = 'INFO'
    )

    $timestamp = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
    $entry = "[$timestamp] [$Level] $Message"

    if ($LogFile) {
        $entry | Out-File -FilePath $LogFile -Append -Encoding utf8
    }

    switch ($Level) {
        'ERROR'   { Write-Host $entry -ForegroundColor Red }
        'WARN'    { Write-Host $entry -ForegroundColor Yellow }
        'SUCCESS' { Write-Host $entry -ForegroundColor Green }
        default   { Write-Verbose $entry }
    }
}

# ─── ADO REST API ──────────────────────────────────────────────────────────────

function Get-AdoAuthHeader {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$Pat
    )

    $base64 = [Convert]::ToBase64String([Text.Encoding]::ASCII.GetBytes(":$Pat"))
    return @{ Authorization = "Basic $base64" }
}

function Invoke-AdoApi {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$Url,
        [Parameter(Mandatory)]
        [string]$Pat,
        [string]$Method = 'GET',
        [object]$Body,
        [string]$ApiVersion = '7.0'
    )

    $separator = if ($Url.Contains('?')) { '&' } else { '?' }
    $fullUrl = "${Url}${separator}api-version=${ApiVersion}"

    $headers = Get-AdoAuthHeader -Pat $Pat
    $headers['Content-Type'] = 'application/json'

    $params = @{
        Uri     = $fullUrl
        Headers = $headers
        Method  = $Method
    }

    if ($Body) {
        $params.Body = ($Body | ConvertTo-Json -Depth 10)
    }

    try {
        $response = Invoke-RestMethod @params -ErrorAction Stop
        return $response
    }
    catch {
        $statusCode = $_.Exception.Response.StatusCode
        throw "ADO API call failed (${statusCode}): $($_.Exception.Message) — URL: $fullUrl"
    }
}

function Get-AdoProjects {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$ServerUrl,
        [Parameter(Mandatory)]
        [string]$Collection,
        [Parameter(Mandatory)]
        [string]$Pat
    )

    $url = "$ServerUrl/$Collection/_apis/projects"
    $result = Invoke-AdoApi -Url $url -Pat $Pat
    return $result.value
}

function Get-TfvcItems {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$ServerUrl,
        [Parameter(Mandatory)]
        [string]$Collection,
        [Parameter(Mandatory)]
        [string]$Pat,
        [string]$ScopePath = '$/',
        [int]$RecursionLevel = 1
    )

    $encodedPath = [Uri]::EscapeDataString($ScopePath)
    $url = "$ServerUrl/$Collection/_apis/tfvc/items?scopePath=${encodedPath}&recursionLevel=${RecursionLevel}"
    $result = Invoke-AdoApi -Url $url -Pat $Pat
    return $result.value
}

function Get-TfvcBranches {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$ServerUrl,
        [Parameter(Mandatory)]
        [string]$Collection,
        [Parameter(Mandatory)]
        [string]$Pat
    )

    $url = "$ServerUrl/$Collection/_apis/tfvc/branches"
    $result = Invoke-AdoApi -Url $url -Pat $Pat
    return $result.value
}

function Get-TfvcChangesets {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$ServerUrl,
        [Parameter(Mandatory)]
        [string]$Collection,
        [Parameter(Mandatory)]
        [string]$Pat,
        [string]$ScopePath,
        [int]$Top = 1
    )

    $url = "$ServerUrl/$Collection/_apis/tfvc/changesets?`$top=$Top"
    if ($ScopePath) {
        $url += "&searchCriteria.itemPath=$([Uri]::EscapeDataString($ScopePath))"
    }
    $result = Invoke-AdoApi -Url $url -Pat $Pat
    return $result.value
}

# ─── Git-TFS Helpers ──────────────────────────────────────────────────────────

function Find-GitTfs {
    [CmdletBinding()]
    param(
        [string]$GitTfsPath
    )

    if ($GitTfsPath -and (Test-Path $GitTfsPath)) {
        return $GitTfsPath
    }

    $found = Get-Command 'git-tfs' -ErrorAction SilentlyContinue
    if ($found) {
        return $found.Source
    }

    $found = Get-Command 'git tfs' -ErrorAction SilentlyContinue
    if ($found) {
        return $found.Source
    }

    throw "git-tfs not found. Install it: https://github.com/git-tfs/git-tfs/releases or set gitTfsPath in config."
}

function Invoke-GitTfs {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$Arguments,
        [string]$WorkingDirectory,
        [string]$LogFile
    )

    $cmd = "git tfs $Arguments"

    Write-MigrationLog -Message "Running: $cmd" -LogFile $LogFile -Level INFO

    $psi = [System.Diagnostics.ProcessStartInfo]::new()
    $psi.FileName = 'git'
    $psi.Arguments = "tfs $Arguments"
    $psi.RedirectStandardOutput = $true
    $psi.RedirectStandardError = $true
    $psi.UseShellExecute = $false
    $psi.CreateNoWindow = $true

    if ($WorkingDirectory) {
        $psi.WorkingDirectory = $WorkingDirectory
    }

    $process = [System.Diagnostics.Process]::Start($psi)
    $stdout = $process.StandardOutput.ReadToEnd()
    $stderr = $process.StandardError.ReadToEnd()
    $process.WaitForExit()

    if ($LogFile) {
        $stdout | Out-File -FilePath $LogFile -Append -Encoding utf8
        if ($stderr) {
            $stderr | Out-File -FilePath $LogFile -Append -Encoding utf8
        }
    }

    if ($process.ExitCode -ne 0) {
        throw "git-tfs failed (exit $($process.ExitCode)): $stderr"
    }

    return $stdout
}

# ─── Author Mapping ──────────────────────────────────────────────────────────

function Read-AuthorMapping {
    [CmdletBinding()]
    param(
        [string]$MappingFile
    )

    if (-not $MappingFile -or -not (Test-Path $MappingFile)) {
        return $null
    }

    $mapping = @{}
    $rows = Import-Csv $MappingFile
    foreach ($row in $rows) {
        $mapping[$row.TfvcIdentity] = "$($row.GitName) <$($row.GitEmail)>"
    }
    return $mapping
}

function Export-AuthorMappingTemplate {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string[]]$TfvcIdentities,
        [Parameter(Mandatory)]
        [string]$OutputPath
    )

    $rows = $TfvcIdentities | Sort-Object -Unique | ForEach-Object {
        [PSCustomObject]@{
            TfvcIdentity = $_
            GitName      = ''
            GitEmail     = ''
        }
    }

    $rows | Export-Csv -Path $OutputPath -NoTypeInformation -Encoding utf8
    return $OutputPath
}

# ─── Git Helpers ──────────────────────────────────────────────────────────────

function Invoke-Git {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$Arguments,
        [string]$WorkingDirectory,
        [string]$LogFile
    )

    $cmd = "git $Arguments"
    Write-MigrationLog -Message "Running: $cmd" -LogFile $LogFile -Level INFO

    $prevDir = $null
    if ($WorkingDirectory) {
        $prevDir = Get-Location
        Set-Location $WorkingDirectory
    }

    try {
        $output = Invoke-Expression $cmd 2>&1
        $exitCode = $LASTEXITCODE

        if ($LogFile) {
            $output | Out-File -FilePath $LogFile -Append -Encoding utf8
        }

        if ($exitCode -ne 0) {
            throw "git failed (exit $exitCode): $output"
        }

        return $output
    }
    finally {
        if ($prevDir) {
            Set-Location $prevDir
        }
    }
}

function Remove-GitTfsMetadata {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$RepoPath,
        [string]$LogFile
    )

    Write-MigrationLog -Message "Removing git-tfs metadata from $RepoPath" -LogFile $LogFile

    # Remove git-tfs remote
    $remotes = Invoke-Git -Arguments "remote" -WorkingDirectory $RepoPath -LogFile $LogFile
    if ($remotes -match 'tfs') {
        Invoke-Git -Arguments "remote remove tfs" -WorkingDirectory $RepoPath -LogFile $LogFile
    }

    # Remove git-tfs-id from commit messages requires filter-repo or filter-branch
    # For now, we leave commit messages as-is since git-tfs-id is informational
    Write-MigrationLog -Message "Note: git-tfs-id lines remain in commit messages (informational only)" -LogFile $LogFile -Level INFO
}

# ─── Cleanup / BFG ───────────────────────────────────────────────────────────

function Remove-LargeFiles {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$RepoPath,
        [string[]]$Extensions,
        [string]$LogFile
    )

    if (-not $Extensions -or $Extensions.Count -eq 0) {
        return
    }

    Write-MigrationLog -Message "Setting up .gitattributes for LFS tracking" -LogFile $LogFile

    Push-Location $RepoPath
    try {
        foreach ($ext in $Extensions) {
            Invoke-Git -Arguments "lfs track `"$ext`"" -WorkingDirectory $RepoPath -LogFile $LogFile
        }
        Invoke-Git -Arguments "add .gitattributes" -WorkingDirectory $RepoPath -LogFile $LogFile
        Invoke-Git -Arguments "commit -m `"Configure Git LFS tracking`"" -WorkingDirectory $RepoPath -LogFile $LogFile
    }
    finally {
        Pop-Location
    }
}

# ─── Validation ──────────────────────────────────────────────────────────────

function Test-AdoConnection {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$ServerUrl,
        [Parameter(Mandatory)]
        [string]$Collection,
        [Parameter(Mandatory)]
        [string]$Pat
    )

    try {
        $projects = Get-AdoProjects -ServerUrl $ServerUrl -Collection $Collection -Pat $Pat
        return @{
            Connected    = $true
            ProjectCount = $projects.Count
            Projects     = $projects.name
        }
    }
    catch {
        return @{
            Connected = $false
            Error     = $_.Exception.Message
        }
    }
}

Export-ModuleMember -Function *
