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

function Get-ConfigAdoServerUrl {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [hashtable]$Config,
        [ValidateSet('Source', 'Target')]
        [string]$Role = 'Source'
    )

    if ($Role -eq 'Target') {
        if ($Config.targetAdoServerUrl) {
            return [string]$Config.targetAdoServerUrl
        }
    }
    else {
        if ($Config.sourceAdoServerUrl) {
            return [string]$Config.sourceAdoServerUrl
        }
    }

    if ($Config.adoServerUrl) {
        return [string]$Config.adoServerUrl
    }

    throw "No ADO server URL configured. Set sourceAdoServerUrl/targetAdoServerUrl (or legacy adoServerUrl)."
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

function Get-AdoBuildDefinitions {
    <#
    .SYNOPSIS
        Gets build/pipeline definitions for a project, optionally filtered by repository path.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$ServerUrl,
        [Parameter(Mandatory)]
        [string]$Collection,
        [Parameter(Mandatory)]
        [string]$Pat,
        [Parameter(Mandatory)]
        [string]$ProjectName
    )

    $encodedProject = [Uri]::EscapeDataString($ProjectName)
    $url = "$ServerUrl/$Collection/$encodedProject/_apis/build/definitions"
    try {
        $result = Invoke-AdoApi -Url $url -Pat $Pat
        return $result.value
    }
    catch {
        return @()
    }
}

# ─── Path Length Mitigation (Windows subst) ───────────────────────────────────

function New-ShortClonePath {
    <#
    .SYNOPSIS
        Creates a very short working path for git-tfs clone operations on Windows.
        Uses subst to map an available drive letter to a short temp directory,
        mitigating the 260-character MAX_PATH limit that git-tfs hits internally.
    .DESCRIPTION
        Even with core.longpaths=true and the registry LongPathsEnabled setting,
        git-tfs may fail because it (and .NET Framework APIs it uses) still respect
        the legacy 260-character limit during checkout operations.
        This function creates a subst drive mapping so the effective working path
        is as short as possible (e.g., "G:\r" instead of "C:\Users\Bob\output\my-repo").
    .OUTPUTS
        Hashtable with: DriveLetter, ShortPath, CleanupScriptBlock
        Returns $null on non-Windows or if subst is unavailable.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$OutputDirectory,
        [string]$LogFile
    )

    # Only applicable on Windows
    if ($env:OS -ne 'Windows_NT') {
        return $null
    }

    # Ensure the output directory exists as the subst target
    if (-not (Test-Path $OutputDirectory)) {
        New-Item -ItemType Directory -Path $OutputDirectory -Force | Out-Null
    }

    # Resolve to absolute path for subst
    $resolvedDir = (Resolve-Path $OutputDirectory).Path

    # Find an available drive letter (prefer letters near end of alphabet to avoid conflicts)
    $preferredLetters = @('G', 'H', 'I', 'J', 'K', 'M', 'N', 'O', 'P', 'Q', 'R', 'S', 'T', 'U', 'V', 'W', 'X', 'Y', 'Z')
    $usedDrives = @((Get-PSDrive -PSProvider FileSystem).Name)

    $driveLetter = $null
    foreach ($letter in $preferredLetters) {
        if ($letter -notin $usedDrives) {
            $driveLetter = $letter
            break
        }
    }

    if (-not $driveLetter) {
        Write-MigrationLog -Message "No available drive letters for subst — skipping path shortening" -LogFile $LogFile -Level WARN
        return $null
    }

    # Create the subst mapping
    try {
        $substResult = & subst "${driveLetter}:" $resolvedDir 2>&1
        if ($LASTEXITCODE -ne 0) {
            Write-MigrationLog -Message "subst failed: $substResult — continuing with original path" -LogFile $LogFile -Level WARN
            return $null
        }
    }
    catch {
        Write-MigrationLog -Message "subst unavailable: $($_.Exception.Message) — continuing with original path" -LogFile $LogFile -Level WARN
        return $null
    }

    Write-MigrationLog -Message "Mapped ${driveLetter}: → $resolvedDir (to shorten file paths for git-tfs)" -LogFile $LogFile -Level INFO

    return @{
        DriveLetter = $driveLetter
        BasePath    = "${driveLetter}:"
        OriginalDir = $resolvedDir
    }
}

function Remove-ShortClonePath {
    <#
    .SYNOPSIS
        Removes a subst drive mapping created by New-ShortClonePath.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$DriveLetter,
        [string]$LogFile
    )

    try {
        & subst "${DriveLetter}:" /d 2>$null
        Write-MigrationLog -Message "Removed subst mapping ${DriveLetter}:" -LogFile $LogFile -Level INFO
    }
    catch {
        Write-MigrationLog -Message "Warning: Could not remove subst ${DriveLetter}: — run 'subst ${DriveLetter}: /d' manually" -LogFile $LogFile -Level WARN
    }
}

function Get-ShortRepoName {
    <#
    .SYNOPSIS
        Generates a short temporary name for clone operations to minimize path length.
        The repo is renamed to its proper name after cloning completes.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$FullRepoName
    )

    # Use a short hash-like name to keep the path minimal during clone
    # Format: "r_" + first 6 chars, keeping total under 10 characters
    $short = $FullRepoName.Substring(0, [Math]::Min(6, $FullRepoName.Length))
    $short = $short -replace '[^a-zA-Z0-9]', ''
    return "r_$short"
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

    throw "git-tfs not found. Install it from https://github.com/git-tfs/git-tfs/releases or set gitTfsPath in your config file."
}

function Invoke-GitTfs {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$Arguments,
        [string]$WorkingDirectory,
        [string]$LogFile,
        [int]$TimeoutMinutes = 0,
        [int]$StallTimeoutMinutes = 0
    )

    $cmd = "git tfs $Arguments"

    # Redact PAT from log output
    $logCmd = $cmd -replace '--password="[^"]*"', '--password="***"'
    Write-MigrationLog -Message "Running: $logCmd" -LogFile $LogFile -Level INFO

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

    # Read stdout/stderr asynchronously to avoid deadlocks
    $stdoutTask = $process.StandardOutput.ReadToEndAsync()
    $stderrTask = $process.StandardError.ReadToEndAsync()

    # Determine effective timeouts
    $hasTimeout = $TimeoutMinutes -gt 0
    $hasStallTimeout = $StallTimeoutMinutes -gt 0
    $timeoutDeadline = if ($hasTimeout) { (Get-Date).AddMinutes($TimeoutMinutes) } else { [datetime]::MaxValue }
    $lastOutputSize = 0
    $lastOutputChangeTime = Get-Date

    # Show a spinner with elapsed time, timeout, and stall detection
    $spinChars = @('|', '/', '-', '\')
    $spinIdx = 0
    $timedOut = $false
    $stalled = $false

    while (-not $process.HasExited) {
        $now = Get-Date
        $elapsed = $now - $process.StartTime
        $spinChar = $spinChars[$spinIdx % 4]
        $spinIdx++

        # Build status line with timeout info
        $statusLine = "  $spinChar  Converting... (elapsed: $($elapsed.ToString('hh\:mm\:ss'))"
        if ($hasTimeout) {
            $remaining = $timeoutDeadline - $now
            if ($remaining.TotalSeconds -gt 0) {
                $statusLine += " | timeout in: $($remaining.ToString('hh\:mm\:ss'))"
            }
        }
        $statusLine += ")  "
        Write-Host "`r$statusLine" -NoNewline -ForegroundColor DarkGray

        # Check for hard timeout
        if ($hasTimeout -and $now -ge $timeoutDeadline) {
            $timedOut = $true
            Write-MigrationLog -Message "TIMEOUT: git-tfs exceeded $TimeoutMinutes minute limit (elapsed: $($elapsed.ToString('hh\:mm\:ss'))). Killing process." -LogFile $LogFile -Level ERROR
            try { $process.Kill() } catch { }
            break
        }

        # Check for stall (no new output for StallTimeoutMinutes)
        if ($hasStallTimeout) {
            $currentSize = $stdoutTask.Result.Length 2>$null
            if ($null -eq $currentSize) { $currentSize = 0 }
            if ($currentSize -ne $lastOutputSize) {
                $lastOutputSize = $currentSize
                $lastOutputChangeTime = $now
            }
            elseif (($now - $lastOutputChangeTime).TotalMinutes -ge $StallTimeoutMinutes) {
                $stalled = $true
                Write-MigrationLog -Message "STALL DETECTED: No new output for $StallTimeoutMinutes minutes. Process appears stuck. Killing." -LogFile $LogFile -Level ERROR
                try { $process.Kill() } catch { }
                break
            }
        }

        Start-Sleep -Milliseconds 500
    }

    if ($timedOut) {
        Write-Host "`r  ✗  TIMED OUT after $TimeoutMinutes minutes.                              " -ForegroundColor Red
        throw "git-tfs timed out after $TimeoutMinutes minutes. The repo may be too large or the server is unresponsive. Skipping this item."
    }

    if ($stalled) {
        Write-Host "`r  ✗  STALLED — no progress for $StallTimeoutMinutes minutes.                " -ForegroundColor Red
        throw "git-tfs stalled (no output for $StallTimeoutMinutes minutes). The process appeared stuck. Skipping this item."
    }

    Write-Host "`r  ✓  Conversion process finished.                              " -ForegroundColor Green

    $stdout = $stdoutTask.GetAwaiter().GetResult()
    $stderr = $stderrTask.GetAwaiter().GetResult()

    if ($LogFile) {
        $stdout | Out-File -FilePath $LogFile -Append -Encoding utf8
        if ($stderr) {
            $stderr | Out-File -FilePath $LogFile -Append -Encoding utf8
        }
    }

    if ($process.ExitCode -ne 0) {
        # Combine stderr and last lines of stdout to surface the actual error
        $detail = ''
        if ($stderr.Trim()) {
            $detail = $stderr.Trim()
        }
        # git-tfs often writes errors to stdout — grab the last few lines
        $stdoutLines = $stdout -split "`n" | Where-Object { $_.Trim() } | Select-Object -Last 10
        if ($stdoutLines) {
            $stdoutTail = ($stdoutLines -join "`n").Trim()
            if ($detail) { $detail += "`n" + $stdoutTail } else { $detail = $stdoutTail }
        }
        if (-not $detail) { $detail = '(no output captured — check the log file)' }
        throw "git-tfs failed (exit $($process.ExitCode)):`n$detail"
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
    $safeCmd = $cmd -replace '://:[^@]+@', '://***@'
    Write-MigrationLog -Message "Running: $safeCmd" -LogFile $LogFile -Level INFO

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
        $friendlyError = Get-FriendlyError -ErrorMessage $_.Exception.Message
        return @{
            Connected = $false
            Error     = $friendlyError
        }
    }
}

function Get-FriendlyError {
    <#
    .SYNOPSIS
        Translates common technical error messages into plain English.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$ErrorMessage
    )

    if ($ErrorMessage -match 'Unauthorized|401') {
        return "Authentication failed — your PAT (Personal Access Token) may be expired or invalid. Generate a new one in ADO and update your config."
    }
    if ($ErrorMessage -match 'Forbidden|403') {
        return "Access denied — your PAT doesn't have the required permissions. Make sure it has 'Code (Read & Write)' and 'Project (Read)' scopes."
    }
    if ($ErrorMessage -match '404|Not Found') {
        return "The server, collection, or project was not found. Double-check the URL and collection name in your config."
    }
    if ($ErrorMessage -match 'Unable to connect|No such host|Name.*not.*resolve|connection.*refused|timed out') {
        return "Cannot reach the ADO server. Check that the server URL in your config is correct and that you're connected to the network (VPN, etc.)."
    }
    if ($ErrorMessage -match '409|already exists') {
        return "A resource with that name already exists at the destination."
    }
    if ($ErrorMessage -match '503|Service Unavailable') {
        return "The ADO server is temporarily unavailable. Wait a moment and try again."
    }
    if ($ErrorMessage -match 'git-tfs.*not found|git tfs.*not recognized') {
        return "The git-tfs tool is not installed. Run Install-Prerequisites.ps1 for installation instructions."
    }
    if ($ErrorMessage -match 'too long|path.*long|248 char|260 char|TF400959|PathTooLong|could not find a part of the path') {
        return "File paths exceed the Windows 260-character limit. " +
            "Try setting outputDirectory to a very short path (e.g. C:\M), " +
            "or migrate specific subfolders instead of the entire repo. " +
            "Deeply nested TFVC paths may need to be renamed in TFS first."
    }

    # Return original if no match
    return $ErrorMessage
}

# ─── Interactive Menu Helpers ─────────────────────────────────────────────────

function Show-MenuHeader {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$Title
    )

    Write-Host ""
    Write-Host "═══════════════════════════════════════════════════════" -ForegroundColor Cyan
    Write-Host "  $Title" -ForegroundColor Cyan
    Write-Host "═══════════════════════════════════════════════════════" -ForegroundColor Cyan
    Write-Host ""
}

function Show-NumberedMenu {
    <#
    .SYNOPSIS
        Displays a numbered list of items and returns the user's selection(s).
    .PARAMETER Items
        Array of display strings.
    .PARAMETER Prompt
        Text shown before the menu.
    .PARAMETER MultiSelect
        Allow comma-separated multiple selections.
    .PARAMETER AllowBack
        Show a "[0] Back" option.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string[]]$Items,
        [string]$Prompt = 'Select an option',
        [switch]$MultiSelect,
        [switch]$AllowBack
    )

    for ($i = 0; $i -lt $Items.Count; $i++) {
        Write-Host "  [$($i + 1)] $($Items[$i])" -ForegroundColor White
    }
    if ($AllowBack) {
        Write-Host "  [0] Back" -ForegroundColor DarkGray
    }
    Write-Host ""

    while ($true) {
        if ($MultiSelect) {
            Write-Host "$Prompt (comma-separated, e.g. 1,3,5): " -ForegroundColor Yellow -NoNewline
        }
        else {
            Write-Host "${Prompt}: " -ForegroundColor Yellow -NoNewline
        }
        $userChoice = Read-Host

        if ($AllowBack -and $userChoice.Trim() -eq '0') {
            return $null
        }

        if ($MultiSelect) {
            $indices = $userChoice -split ',' | ForEach-Object {
                $val = $_.Trim()
                if ($val -match '^\d+$') { [int]$val - 1 }
            }
            $valid = $indices | Where-Object { $_ -ge 0 -and $_ -lt $Items.Count }
            if ($valid.Count -gt 0) {
                return $valid
            }
        }
        else {
            if ($userChoice.Trim() -match '^\d+$') {
                $idx = [int]$userChoice.Trim() - 1
                if ($idx -ge 0 -and $idx -lt $Items.Count) {
                    return $idx
                }
            }
        }

        Write-Host "  Invalid selection. Please enter a number from the list above." -ForegroundColor Red
    }
}

function Select-AdoCollection {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [hashtable]$Config,
        [string]$Prompt = 'Select a collection'
    )

    $collectionNames = @($Config.collections.Keys | Sort-Object)
    if ($collectionNames.Count -eq 0) {
        throw "No collections defined in config."
    }

    $displayItems = $collectionNames | ForEach-Object {
        $desc = $Config.collections[$_].description
        if ($desc) { "$_ — $desc" } else { $_ }
    }

    Show-MenuHeader -Title $Prompt
    $idx = Show-NumberedMenu -Items $displayItems -Prompt $Prompt
    if ($null -eq $idx) { return $null }

    return $collectionNames[$idx]
}

function Select-AdoProject {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [hashtable]$Config,
        [Parameter(Mandatory)]
        [string]$Collection,
        [string]$ServerUrl,
        [string]$Prompt = 'Select a project'
    )

    $pat = $Config.collections[$Collection].pat
    if (-not $ServerUrl) {
        $ServerUrl = Get-ConfigAdoServerUrl -Config $Config -Role Source
    }
    Write-Host "  Fetching projects from '$Collection'..." -ForegroundColor DarkGray
    $projects = Get-AdoProjects -ServerUrl $ServerUrl -Collection $Collection -Pat $pat
    $projectNames = @($projects.name | Sort-Object)

    if ($projectNames.Count -eq 0) {
        Write-Host "  No projects found in collection '$Collection'." -ForegroundColor Red
        return $null
    }

    Show-MenuHeader -Title "$Prompt (Collection: $Collection)"
    $idx = Show-NumberedMenu -Items $projectNames -Prompt $Prompt -AllowBack
    if ($null -eq $idx) { return $null }

    return $projectNames[$idx]
}

function Select-TfvcFolders {
    <#
    .SYNOPSIS
        Lists TFVC folders under a given path and lets the user pick one or more.
    .OUTPUTS
        Array of selected TFVC folder paths, or $null if user chose Back.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [hashtable]$Config,
        [Parameter(Mandatory)]
        [string]$Collection,
        [Parameter(Mandatory)]
        [string]$ProjectName,
        [string]$ServerUrl,
        [string]$ParentPath,
        [switch]$MultiSelect,
        [string]$Prompt = 'Select folder(s)'
    )

    $pat = $Config.collections[$Collection].pat
    if (-not $ServerUrl) {
        $ServerUrl = Get-ConfigAdoServerUrl -Config $Config -Role Source
    }
    $rootPath = if ($ParentPath) { $ParentPath } else { "`$/$ProjectName" }

    Write-Host "  Fetching TFVC contents under '$rootPath'..." -ForegroundColor DarkGray

    try {
        $items = Get-TfvcItems -ServerUrl $ServerUrl -Collection $Collection `
            -Pat $pat -ScopePath $rootPath -RecursionLevel 1
    }
    catch {
        Write-Host "  Could not list items under '$rootPath': $($_.Exception.Message)" -ForegroundColor Red
        return $null
    }

    $folders = @($items | Where-Object { $_.isFolder -eq $true -and $_.path -ne $rootPath } | Sort-Object path)

    if ($folders.Count -eq 0) {
        Write-Host "  No subfolders found under '$rootPath'." -ForegroundColor Yellow
        return $null
    }

    # Build display with last-changeset info
    $displayItems = foreach ($folder in $folders) {
        $lastInfo = ''
        try {
            $cs = Get-TfvcChangesets -ServerUrl $ServerUrl -Collection $Collection `
                -Pat $pat -ScopePath $folder.path -Top 1
            if ($cs) {
                $date = ([datetime]$cs[0].createdDate).ToString('yyyy-MM-dd')
                $lastInfo = "  (last change: $date)"
            }
        }
        catch { }
        "$($folder.path)$lastInfo"
    }

    Show-MenuHeader -Title "$Prompt (under $rootPath)"
    $selectedIndices = Show-NumberedMenu -Items $displayItems -Prompt $Prompt -MultiSelect:$MultiSelect -AllowBack

    if ($null -eq $selectedIndices) { return $null }

    if ($MultiSelect) {
        return @($selectedIndices | ForEach-Object { $folders[$_].path })
    }
    else {
        return @($folders[$selectedIndices].path)
    }
}

Export-ModuleMember -Function *
