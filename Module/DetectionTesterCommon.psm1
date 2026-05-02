<#
.SYNOPSIS
    Shared module for Detection Method Testing Tool.

.DESCRIPTION
    Provides:
      - Structured logging (Initialize-Logging, Write-Log)
      - Detection tests (Test-RegistryKeyValueDetection, Test-RegistryKeyDetection,
        Test-FileDetection, Test-ScriptDetection, Test-CompoundDetection)
      - ARP enumeration (Get-InstalledApplications)
      - Manifest import (Import-DetectionManifest)
      - Export to CSV and HTML (Export-DetectionResultsCsv, Export-DetectionResultsHtml)

.EXAMPLE
    Import-Module "$PSScriptRoot\Module\DetectionTesterCommon.psd1" -Force
    Initialize-Logging -LogPath "C:\projects\detectiontester\Logs\test.log"
#>

# ---------------------------------------------------------------------------
# Module-scoped state
# ---------------------------------------------------------------------------

$script:__LogPath = $null

# ---------------------------------------------------------------------------
# Logging
# ---------------------------------------------------------------------------

function Initialize-Logging {
    param([string]$LogPath)

    $script:__LogPath = $LogPath

    if ($LogPath) {
        $parentDir = Split-Path -Path $LogPath -Parent
        if ($parentDir -and -not (Test-Path -LiteralPath $parentDir)) {
            New-Item -ItemType Directory -Path $parentDir -Force | Out-Null
        }

        $header = "[{0}] [INFO ] === Log initialized ===" -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss')
        Set-Content -LiteralPath $LogPath -Value $header -Encoding UTF8
    }
}

function Write-Log {
    <#
    .SYNOPSIS
        Writes a timestamped, severity-tagged log message.
    #>
    param(
        [AllowEmptyString()]
        [Parameter(Mandatory, Position = 0)]
        [string]$Message,

        [ValidateSet('INFO', 'WARN', 'ERROR')]
        [string]$Level = 'INFO',

        [switch]$Quiet
    )

    $timestamp = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
    $formatted = "[{0}] [{1,-5}] {2}" -f $timestamp, $Level, $Message

    if (-not $Quiet) {
        Write-Host $formatted
        if ($Level -eq 'ERROR') {
            $host.UI.WriteErrorLine($formatted)
        }
    }

    if ($script:__LogPath) {
        Add-Content -LiteralPath $script:__LogPath -Value $formatted -Encoding UTF8 -ErrorAction SilentlyContinue
    }
}

# ---------------------------------------------------------------------------
# Private helper: value comparison matching MECM behavior
# ---------------------------------------------------------------------------

function Compare-DetectionValue {
    <#
    .SYNOPSIS
        Compares two values using MECM detection method semantics.
    .DESCRIPTION
        PropertyType chooses the comparison kind:
          - 'Version' (default for greater/less): parse as System.Version (1.2.3.4 form)
          - 'Integer' / 'DWORD' / 'QWORD': parse as Int64
          - 'String' (default): culture-invariant string compare
        IsEquals always falls back to string equality if numeric parse fails.
    .OUTPUTS
        [bool] True if comparison succeeds.
    #>
    param(
        [AllowEmptyString()][AllowNull()][string]$Actual,
        [AllowEmptyString()][AllowNull()][string]$Expected,
        [string]$Operator = 'IsEquals',
        [string]$PropertyType = 'String'
    )

    if ($null -eq $Actual) { return $false }

    if ($Operator -eq 'IsEquals') { return ($Actual -eq $Expected) }

    # For greater/less, choose numeric kind by PropertyType.
    $isInteger = $PropertyType -in @('Integer','Int','Int32','Int64','DWORD','QWORD')
    $isVersion = $PropertyType -in @('Version','FileVersion','ProductVersion','String')

    if ($isInteger) {
        try {
            $iActual   = [int64]$Actual
            $iExpected = [int64]$Expected
            switch ($Operator) {
                'GreaterEquals' { return ($iActual -ge $iExpected) }
                'GreaterThan'   { return ($iActual -gt $iExpected) }
                'LessEquals'    { return ($iActual -le $iExpected) }
                'LessThan'      { return ($iActual -lt $iExpected) }
            }
        } catch {
            Write-Log "Integer parse failed: Actual='$Actual', Expected='$Expected' - $_" -Level WARN -Quiet
            return $false
        }
    }

    if ($isVersion) {
        try {
            $vActual   = [System.Version]$Actual
            $vExpected = [System.Version]$Expected
            switch ($Operator) {
                'GreaterEquals' { return ($vActual -ge $vExpected) }
                'GreaterThan'   { return ($vActual -gt $vExpected) }
                'LessEquals'    { return ($vActual -le $vExpected) }
                'LessThan'      { return ($vActual -lt $vExpected) }
            }
        } catch {
            # Fall through to integer attempt below for ambiguous cases (e.g. "8246" parses as a version OR an int).
            Write-Log "Compare-DetectionValue: Version parse failed for Actual='$Actual' Expected='$Expected'; trying Int64 path." -Level WARN -Quiet
        }
        try {
            $iActual   = [int64]$Actual
            $iExpected = [int64]$Expected
            switch ($Operator) {
                'GreaterEquals' { return ($iActual -ge $iExpected) }
                'GreaterThan'   { return ($iActual -gt $iExpected) }
                'LessEquals'    { return ($iActual -le $iExpected) }
                'LessThan'      { return ($iActual -lt $iExpected) }
            }
        } catch {
            Write-Log "Version/Integer parse failed: Actual='$Actual', Expected='$Expected' - $_" -Level WARN -Quiet
            return $false
        }
    }

    Write-Log "Unknown operator/PropertyType combo: Operator=$Operator PropertyType=$PropertyType" -Level WARN -Quiet
    return $false
}

# ---------------------------------------------------------------------------
# Detection tests
# ---------------------------------------------------------------------------

function Get-DTHiveDisplayPrefix {
    <#
    .SYNOPSIS
        Returns the display prefix ('HKLM' or 'HKCU') for a hive name.
    #>
    param(
        [ValidateSet('LocalMachine','CurrentUser')]
        [string]$Hive
    )
    if ($Hive -eq 'CurrentUser') { 'HKCU' } else { 'HKLM' }
}

function ConvertTo-DTHiveName {
    <#
    .SYNOPSIS
        Normalizes hive aliases (HKLM, HKCU, HKEY_LOCAL_MACHINE, etc.) to
        the canonical .NET names ('LocalMachine' or 'CurrentUser').
    .OUTPUTS
        [string] one of 'LocalMachine' or 'CurrentUser'.
    #>
    param([string]$Value)
    if ([string]::IsNullOrWhiteSpace($Value)) { return 'LocalMachine' }
    switch -Regex ($Value.Trim()) {
        '^(HKCU|HKEY_CURRENT_USER|CurrentUser)$' { return 'CurrentUser' }
        '^(HKLM|HKEY_LOCAL_MACHINE|LocalMachine)$' { return 'LocalMachine' }
        default { return 'LocalMachine' }
    }
}

function ConvertTo-DTRegistryRelativePath {
    <#
    .SYNOPSIS
        Strips a hive prefix from a registry path so OpenSubKey can use it
        as a relative path. Handles all common prefix forms:
          HKLM\..., HKCU\...
          HKEY_LOCAL_MACHINE\..., HKEY_CURRENT_USER\...
          HKLM:\..., HKCU:\...                (PowerShell PSDrive form)
          Registry::HKLM\..., Registry::HKEY_CURRENT_USER\...
          Microsoft.PowerShell.Core\Registry::HKLM\...
    .OUTPUTS
        [string] hive-stripped path; empty input returns empty string.
    #>
    param([string]$Path)
    if ([string]::IsNullOrWhiteSpace($Path)) { return '' }
    $p = $Path.Trim()
    # Strip the PowerShell registry provider prefix if present.
    $p = $p -replace '^(Registry::|Microsoft\.PowerShell\.Core\\Registry::)', ''
    # Strip any hive name with either a backslash or a colon-then-backslash separator.
    $p = $p -replace '^(HKLM|HKCU|HKEY_LOCAL_MACHINE|HKEY_CURRENT_USER)(:\\|\\|:)', ''
    return $p.TrimStart('\')
}

function Get-DTHiveFromPath {
    <#
    .SYNOPSIS
        Inspects a registry path string and returns the hive it begins
        with ('LocalMachine', 'CurrentUser', or '' when the path has no
        hive prefix). Recognizes all the prefix forms that
        ConvertTo-DTRegistryRelativePath strips.
    .OUTPUTS
        [string] one of 'LocalMachine', 'CurrentUser', or '' (empty).
    #>
    param([string]$Path)
    if ([string]::IsNullOrWhiteSpace($Path)) { return '' }
    $p = $Path.Trim()
    # Drop any provider prefix so we can inspect the hive token directly.
    $p = $p -replace '^(Registry::|Microsoft\.PowerShell\.Core\\Registry::)', ''
    if ($p -match '^(HKCU|HKEY_CURRENT_USER)(:\\|\\|:|$)')  { return 'CurrentUser'  }
    if ($p -match '^(HKLM|HKEY_LOCAL_MACHINE)(:\\|\\|:|$)') { return 'LocalMachine' }
    return ''
}

function Test-RegistryKeyValueDetection {
    <#
    .SYNOPSIS
        Tests a RegistryKeyValue detection rule against the local machine.
    .OUTPUTS
        [hashtable] with keys: Detected, KeyExists, ValueFound, ActualValue, ExpectedValue, Operator, Details
    #>
    param(
        [Parameter(Mandatory)][string]$RegistryKeyRelative,
        [Parameter(Mandatory)][string]$ValueName,
        [string]$ExpectedValue,
        [string]$Operator = 'IsEquals',
        [string]$PropertyType = 'String',
        [bool]$Is64Bit = $true,
        [ValidateSet('LocalMachine','CurrentUser')]
        [string]$Hive = 'LocalMachine'
    )

    $hivePrefix = Get-DTHiveDisplayPrefix -Hive $Hive
    $result = @{
        Type          = 'RegistryKeyValue'
        Target        = "${hivePrefix}\$RegistryKeyRelative"
        Detected      = $false
        KeyExists     = $false
        ValueFound    = $false
        ActualValue   = $null
        ExpectedValue = $ExpectedValue
        Operator      = $Operator
        Hive          = $Hive
        Details       = ''
    }

    try {
        $view = if ($Is64Bit) {
            [Microsoft.Win32.RegistryView]::Registry64
        } else {
            [Microsoft.Win32.RegistryView]::Registry32
        }

        $hiveEnum = if ($Hive -eq 'CurrentUser') {
            [Microsoft.Win32.RegistryHive]::CurrentUser
        } else {
            [Microsoft.Win32.RegistryHive]::LocalMachine
        }

        $rootKey = [Microsoft.Win32.RegistryKey]::OpenBaseKey($hiveEnum, $view)
        $subKey  = $rootKey.OpenSubKey($RegistryKeyRelative)

        if ($null -eq $subKey) {
            $rootKey.Close()
            $result.Details = "Key not found: ${hivePrefix}\$RegistryKeyRelative"
            Write-Log "RegistryKeyValue: Key not found - ${hivePrefix}\$RegistryKeyRelative" -Quiet
            return $result
        }

        $result.KeyExists = $true
        $rawValue = $subKey.GetValue($ValueName, $null)
        $subKey.Close()
        $rootKey.Close()

        if ($null -eq $rawValue) {
            $result.Details = "Value '$ValueName' not found in key"
            Write-Log "RegistryKeyValue: Value '$ValueName' not found" -Quiet
            return $result
        }

        $result.ValueFound = $true
        $result.ActualValue = [string]$rawValue

        $match = Compare-DetectionValue -Actual ([string]$rawValue) -Expected $ExpectedValue -Operator $Operator -PropertyType $PropertyType
        $result.Detected = $match

        if ($match) {
            $result.Details = "MATCH ($Operator): '$($result.ActualValue)' vs '$ExpectedValue'"
        } else {
            $result.Details = "NO MATCH ($Operator): '$($result.ActualValue)' vs '$ExpectedValue'"
        }

        Write-Log "RegistryKeyValue: $($result.Details)" -Quiet
    } catch {
        $result.Details = "Error: $_"
        Write-Log "RegistryKeyValue error: $_" -Level ERROR -Quiet
    }

    return $result
}

function Test-RegistryKeyDetection {
    <#
    .SYNOPSIS
        Tests a RegistryKey existence detection rule against the local machine.
    .OUTPUTS
        [hashtable] with keys: Detected, KeyExists, Details
    #>
    param(
        [Parameter(Mandatory)][string]$RegistryKeyRelative,
        [bool]$Is64Bit = $true,
        [ValidateSet('LocalMachine','CurrentUser')]
        [string]$Hive = 'LocalMachine'
    )

    $hivePrefix = Get-DTHiveDisplayPrefix -Hive $Hive
    $result = @{
        Type          = 'RegistryKey'
        Target        = "${hivePrefix}\$RegistryKeyRelative"
        Detected      = $false
        KeyExists     = $false
        ActualValue   = $null
        ExpectedValue = '(key exists)'
        Operator      = 'Existence'
        Hive          = $Hive
        Details       = ''
    }

    try {
        $view = if ($Is64Bit) {
            [Microsoft.Win32.RegistryView]::Registry64
        } else {
            [Microsoft.Win32.RegistryView]::Registry32
        }

        $hiveEnum = if ($Hive -eq 'CurrentUser') {
            [Microsoft.Win32.RegistryHive]::CurrentUser
        } else {
            [Microsoft.Win32.RegistryHive]::LocalMachine
        }

        $rootKey = [Microsoft.Win32.RegistryKey]::OpenBaseKey($hiveEnum, $view)
        try {
            $subKey  = $rootKey.OpenSubKey($RegistryKeyRelative)
            if ($null -ne $subKey) {
                $result.Detected = $true
                $result.KeyExists = $true
                $result.Details = "Key exists: ${hivePrefix}\$RegistryKeyRelative"
                $subKey.Close()
            } else {
                $result.Details = "Key not found: ${hivePrefix}\$RegistryKeyRelative"
            }
        } finally {
            $rootKey.Close()
        }

        Write-Log "RegistryKey: $($result.Details)" -Quiet
    } catch {
        $result.Details = "Error: $_"
        Write-Log "RegistryKey error: $_" -Level ERROR -Quiet
    }

    return $result
}

function Test-FileDetection {
    <#
    .SYNOPSIS
        Tests a File detection rule against the local machine.
    .OUTPUTS
        [hashtable] with keys: Detected, FileExists, ActualValue, ExpectedValue, Operator, Details
    #>
    param(
        [Parameter(Mandatory)][string]$FilePath,
        [Parameter(Mandatory)][string]$FileName,
        [string]$PropertyType = 'Existence',
        [string]$ExpectedValue,
        [string]$Operator = 'GreaterEquals',
        [bool]$Is64Bit = $true
    )

    $result = @{
        Type          = 'File'
        Target        = (Join-Path $FilePath $FileName)
        Detected      = $false
        FileExists    = $false
        ActualValue   = $null
        ExpectedValue = if ($PropertyType -eq 'Existence') { '(file exists)' } else { $ExpectedValue }
        Operator      = if ($PropertyType -eq 'Existence') { 'Existence' } else { $Operator }
        Details       = ''
    }

    try {
        $expandedPath = [System.Environment]::ExpandEnvironmentVariables($FilePath)
        # Approximate WOW64 file-system redirection: when Is64Bit=$false on a
        # 64-bit OS, the MECM agent runs detection in 32-bit mode and
        # %SystemRoot%\System32 transparently redirects to SysWOW64. We're a
        # 64-bit PS host, so simulate the redirection by string-rewriting.
        if (-not $Is64Bit -and [Environment]::Is64BitOperatingSystem) {
            $expandedPath = $expandedPath -ireplace '\\System32(\\|$)', '\SysWOW64$1'
        }
        $fullPath = Join-Path $expandedPath $FileName

        if (-not (Test-Path -LiteralPath $fullPath)) {
            $result.Details = "File not found: $fullPath"
            Write-Log "File: $($result.Details)" -Quiet
            return $result
        }

        $result.FileExists = $true

        if ($PropertyType -eq 'Existence') {
            $result.Detected = $true
            $result.Details = "File exists: $fullPath"
        } elseif ($PropertyType -eq 'Version') {
            $versionInfo = [System.Diagnostics.FileVersionInfo]::GetVersionInfo($fullPath)
            $fileVersion = $versionInfo.FileVersion

            if ([string]::IsNullOrWhiteSpace($fileVersion)) {
                $result.Details = "File exists but has no version information"
                Write-Log "File: No version info for $fullPath" -Quiet
                return $result
            }

            # Clean version string: strip build metadata, parenthetical suffixes, spaces
            $fileVersion = $fileVersion.Trim() -replace '\s*\(.*\)\s*$', '' -replace '\s+', ''
            $result.ActualValue = $fileVersion

            $match = Compare-DetectionValue -Actual $fileVersion -Expected $ExpectedValue -Operator $Operator -PropertyType 'Version'
            $result.Detected = $match

            if ($match) {
                $result.Details = "MATCH ($Operator): version '$fileVersion' vs '$ExpectedValue'"
            } else {
                $result.Details = "NO MATCH ($Operator): version '$fileVersion' vs '$ExpectedValue'"
            }
        } else {
            $result.Details = "Unsupported PropertyType: $PropertyType"
        }

        Write-Log "File: $($result.Details)" -Quiet
    } catch {
        $result.Details = "Error: $_"
        Write-Log "File error: $_" -Level ERROR -Quiet
    }

    return $result
}

function Test-ScriptDetection {
    <#
    .SYNOPSIS
        Tests a Script detection rule by executing a PowerShell scriptblock.
    .DESCRIPTION
        Any non-empty stdout from the script means DETECTED (matches MECM behavior).
        The script runs in an isolated runspace so 'exit' / 'throw' don't kill
        the host, and a timeout (default 30s) prevents an infinite loop or
        Start-Sleep from freezing the GUI permanently.

        ISOLATION LIMITATION: process-level APIs ([Environment]::Exit,
        [Process]::Kill on the parent PID, kernel32!ExitProcess) cannot be
        contained by an in-process runspace. Detection scripts under test
        are assumed to be packager-authored code, not adversarial input.
        For full isolation against arbitrary scripts, drive a child
        powershell.exe process - that's out of scope for this tool's
        v1.0 use case (test bench for trusted MECM detection rules).
    .OUTPUTS
        [hashtable] with keys: Detected, ScriptOutput, Details, TimedOut
    #>
    param(
        [Parameter(Mandatory)][string]$ScriptText,
        [string]$ScriptLanguage = 'PowerShell',
        [int]$TimeoutSeconds = 30
    )

    $result = @{
        Type          = 'Script'
        Target        = '(script)'
        Detected      = $false
        ActualValue   = $null
        ExpectedValue = '(non-empty output)'
        Operator      = 'Script'
        ScriptOutput  = $null
        TimedOut      = $false
        Details       = ''
    }

    if ($ScriptLanguage -ne 'PowerShell') {
        $result.Details = "Unsupported script language: $ScriptLanguage (only PowerShell supported)"
        Write-Log "Script: $($result.Details)" -Level WARN -Quiet
        return $result
    }

    $rs = $null
    $ps = $null
    try {
        $rs = [runspacefactory]::CreateRunspace()
        $rs.ApartmentState = 'STA'
        $rs.Open()
        $ps = [powershell]::Create()
        $ps.Runspace = $rs
        [void]$ps.AddScript($ScriptText)

        $async = $ps.BeginInvoke()
        $deadline = [DateTime]::UtcNow.AddSeconds([Math]::Max(1, $TimeoutSeconds))
        while (-not $async.IsCompleted -and [DateTime]::UtcNow -lt $deadline) {
            Start-Sleep -Milliseconds 100
        }

        if (-not $async.IsCompleted) {
            try { $ps.Stop() } catch { Write-Log "Test-ScriptDetection: Stop() on timeout swallowed: $($_.Exception.Message)" -Level WARN -Quiet }
            $result.TimedOut = $true
            $result.Details = "Script timed out after ${TimeoutSeconds}s (no output captured); treated as NOT DETECTED."
            Write-Log "Script: $($result.Details)" -Level WARN -Quiet
            return $result
        }

        $output = $ps.EndInvoke($async)
        $stdout = @($output | ForEach-Object { [string]$_ }) -join "`r`n"
        $stderr = @($ps.Streams.Error | ForEach-Object { [string]$_ }) -join "`r`n"

        $result.ScriptOutput = $stdout
        $result.ActualValue  = $stdout

        if (-not [string]::IsNullOrWhiteSpace($stdout)) {
            $result.Detected = $true
            $preview = $stdout.Substring(0, [Math]::Min(200, $stdout.Length))
            $result.Details = "Script produced output (DETECTED): $preview"
        } else {
            $result.Details = "Script produced no output (NOT DETECTED)"
            if (-not [string]::IsNullOrWhiteSpace($stderr)) {
                $errPreview = $stderr.Substring(0, [Math]::Min(200, $stderr.Length))
                $result.Details += " | Errors: $errPreview"
            }
        }

        Write-Log "Script: $($result.Details)" -Quiet
    } catch {
        $result.Details = "Script execution error: $_"
        Write-Log "Script error: $_" -Level ERROR -Quiet
    } finally {
        # Dispose failures can't be acted on; log and continue so cleanup races don't leak.
        if ($ps) { try { $ps.Dispose() } catch { Write-Log "Test-ScriptDetection: ps.Dispose swallowed" -Level WARN -Quiet } }
        if ($rs) { try { $rs.Close(); $rs.Dispose() } catch { Write-Log "Test-ScriptDetection: rs cleanup swallowed" -Level WARN -Quiet } }
    }

    return $result
}

function Test-CompoundDetection {
    <#
    .SYNOPSIS
        Tests a compound detection rule (And/Or) by evaluating each clause.
    .OUTPUTS
        [hashtable] with keys: Detected, Connector, ClauseResults, Details
    #>
    param(
        [Parameter(Mandatory)][string]$Connector,
        [Parameter(Mandatory)][array]$Clauses
    )

    $result = @{
        Type           = 'Compound'
        Target         = "(${Connector}: $($Clauses.Count) clauses)"
        Detected       = $false
        ActualValue    = $null
        ExpectedValue  = "(all $Connector)"
        Operator       = $Connector
        ClauseResults  = @()
        Details        = ''
    }

    $clauseResults = @()

    foreach ($clause in $Clauses) {
        $type = if ($clause.Type) { $clause.Type } else { 'RegistryKeyValue' }

        $clauseResult = switch ($type) {
            'RegistryKeyValue' {
                $p = @{
                    RegistryKeyRelative = $clause.RegistryKeyRelative
                    ValueName           = if ($clause.ValueName) { $clause.ValueName } else { 'DisplayVersion' }
                    ExpectedValue       = if ($clause.ExpectedValue) { $clause.ExpectedValue } else { $clause.DisplayVersion }
                    Operator            = if ($clause.Operator) { $clause.Operator } else { 'IsEquals' }
                    PropertyType        = if ($clause.PropertyType) { $clause.PropertyType } else { 'String' }
                    Is64Bit             = if ($null -ne $clause.Is64Bit) { [bool]$clause.Is64Bit } else { $true }
                    Hive                = if ($clause.Hive) { [string]$clause.Hive } else { 'LocalMachine' }
                }
                Test-RegistryKeyValueDetection @p
            }
            'RegistryKey' {
                $p = @{
                    RegistryKeyRelative = $clause.RegistryKeyRelative
                    Is64Bit             = if ($null -ne $clause.Is64Bit) { [bool]$clause.Is64Bit } else { $true }
                    Hive                = if ($clause.Hive) { [string]$clause.Hive } else { 'LocalMachine' }
                }
                Test-RegistryKeyDetection @p
            }
            'File' {
                $p = @{
                    FilePath      = $clause.FilePath
                    FileName      = $clause.FileName
                    PropertyType  = if ($clause.PropertyType) { $clause.PropertyType } else { 'Existence' }
                    ExpectedValue = $clause.ExpectedValue
                    Operator      = if ($clause.Operator) { $clause.Operator } else { 'GreaterEquals' }
                    Is64Bit       = if ($null -ne $clause.Is64Bit) { [bool]$clause.Is64Bit } else { $true }
                }
                Test-FileDetection @p
            }
            'Script' {
                Test-ScriptDetection -ScriptText $clause.ScriptText -ScriptLanguage $(if ($clause.ScriptLanguage) { $clause.ScriptLanguage } else { 'PowerShell' })
            }
            default {
                @{ Detected = $false; Details = "Unknown clause type: $type" }
            }
        }

        $clauseResults += $clauseResult
    }

    $result.ClauseResults = $clauseResults
    $detected = $clauseResults | ForEach-Object { $_.Detected }

    if ($Connector -eq 'And') {
        $result.Detected = ($detected -notcontains $false)
    } elseif ($Connector -eq 'Or') {
        $result.Detected = ($detected -contains $true)
    }

    $passCount = ($detected | Where-Object { $_ -eq $true }).Count
    $result.Details = "${Connector}: $passCount/$($Clauses.Count) clauses passed - $(if ($result.Detected) { 'DETECTED' } else { 'NOT DETECTED' })"
    $result.ActualValue = "$passCount/$($Clauses.Count) passed"

    Write-Log "Compound: $($result.Details)" -Quiet
    return $result
}

# ---------------------------------------------------------------------------
# ARP enumeration
# ---------------------------------------------------------------------------

function Get-InstalledApplications {
    <#
    .SYNOPSIS
        Enumerates installed applications from both ARP registry hives.
    .OUTPUTS
        [array] of [PSCustomObject] with DisplayName, Publisher, DisplayVersion, Architecture, RegistryKey, UninstallString
    #>

    $apps = @()
    # HKLM (machine-wide) and HKCU (per-user) ARP paths.
    # HKCU has only one path - per-user installs aren't architecture-split.
    $arpPaths = @(
        @{ Hive = 'LocalMachine'; Path = 'SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall';            Arch = 'x64'; Scope = 'Machine' }
        @{ Hive = 'LocalMachine'; Path = 'SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall'; Arch = 'x86'; Scope = 'Machine' }
        @{ Hive = 'CurrentUser';  Path = 'SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall';            Arch = 'x64'; Scope = 'User' }
    )

    foreach ($arp in $arpPaths) {
        try {
            $hiveEnum = if ($arp.Hive -eq 'CurrentUser') {
                [Microsoft.Win32.RegistryHive]::CurrentUser
            } else {
                [Microsoft.Win32.RegistryHive]::LocalMachine
            }

            $rootKey = [Microsoft.Win32.RegistryKey]::OpenBaseKey(
                $hiveEnum,
                [Microsoft.Win32.RegistryView]::Registry64
            )
            $parentKey = $rootKey.OpenSubKey($arp.Path)
            if ($null -eq $parentKey) {
                $rootKey.Close()
                continue
            }

            $hivePrefix = Get-DTHiveDisplayPrefix -Hive $arp.Hive

            foreach ($subKeyName in $parentKey.GetSubKeyNames()) {
                try {
                    $subKey = $parentKey.OpenSubKey($subKeyName)
                    if ($null -eq $subKey) { continue }

                    $displayName = $subKey.GetValue('DisplayName', $null)
                    if ([string]::IsNullOrWhiteSpace($displayName)) {
                        $subKey.Close()
                        continue
                    }

                    $apps += [PSCustomObject]@{
                        DisplayName          = [string]$displayName
                        Publisher            = [string]$subKey.GetValue('Publisher', '')
                        DisplayVersion       = [string]$subKey.GetValue('DisplayVersion', '')
                        Architecture         = $arp.Arch
                        Scope                = $arp.Scope
                        Hive                 = $arp.Hive
                        RegistryKey          = "${hivePrefix}\$($arp.Path)\$subKeyName"
                        RegistryKeyRelative  = "$($arp.Path)\$subKeyName"
                        UninstallString      = [string]$subKey.GetValue('UninstallString', '')
                        QuietUninstallString = [string]$subKey.GetValue('QuietUninstallString', '')
                        InstallLocation      = [string]$subKey.GetValue('InstallLocation', '')
                        InstallDate          = [string]$subKey.GetValue('InstallDate', '')
                    }

                    $subKey.Close()
                } catch {
                    # Access denied on individual ARP subkeys is normal; skip and continue.
                    Write-Log "Get-InstalledApplications: skipped subkey '$subKeyName' ($($_.Exception.Message))" -Level WARN -Quiet
                }
            }

            $parentKey.Close()
            $rootKey.Close()
        } catch {
            Write-Log "ARP enumeration error ($($arp.Hive) $($arp.Arch)): $_" -Level WARN -Quiet
        }
    }

    Write-Log ("Enumerated {0} installed applications (HKLM + HKCU)" -f $apps.Count) -Quiet
    return ($apps | Sort-Object DisplayName)
}

# ---------------------------------------------------------------------------
# Manifest import
# ---------------------------------------------------------------------------

function Import-DetectionManifest {
    <#
    .SYNOPSIS
        Parses a stage-manifest.json from the Application Packager and returns the Detection block.
    .DESCRIPTION
        Normalizes field names to handle variations between manifest formats.
    .OUTPUTS
        [hashtable] with normalized detection fields.
    #>
    param(
        [Parameter(Mandatory)][string]$Path
    )

    if (-not (Test-Path -LiteralPath $Path)) {
        throw "Manifest file not found: $Path"
    }

    try {
        $manifest = Get-Content -LiteralPath $Path -Raw | ConvertFrom-Json
    } catch {
        throw "Failed to parse manifest JSON: $_"
    }

    $det = $manifest.Detection
    if ($null -eq $det) {
        throw "No 'Detection' block found in manifest"
    }

    # Handle compound detection (array of clauses with connector)
    if ($det.Clauses -and $det.Connector) {
        $normalizedClauses = @()
        foreach ($clause in $det.Clauses) {
            $normalizedClauses += (Resolve-DetectionFields -Det $clause)
        }
        return @{
            Type      = 'Compound'
            Connector = [string]$det.Connector
            Clauses   = $normalizedClauses
        }
    }

    return (Resolve-DetectionFields -Det $det)
}

function Resolve-DetectionFields {
    <#
    .SYNOPSIS
        Normalizes a single detection block's field names.
    #>
    param([Parameter(Mandatory)]$Det)

    $type = if ($Det.Type) { [string]$Det.Type } else { 'RegistryKeyValue' }

    $normalized = @{ Type = $type }

    switch ($type) {
        'RegistryKeyValue' {
            $rawPath = if ($Det.RegistryKeyRelative) { [string]$Det.RegistryKeyRelative } elseif ($Det.KeyPath) { [string]$Det.KeyPath } else { '' }
            $rawHive = if ($Det.Hive) { [string]$Det.Hive } else { '' }
            # If hive missing, infer from the path (covers all prefix forms incl. HKCU:\ and Registry::HKEY_*\).
            if (-not $rawHive) { $rawHive = Get-DTHiveFromPath -Path $rawPath }
            $normalized.RegistryKeyRelative = ConvertTo-DTRegistryRelativePath -Path $rawPath
            $normalized.ValueName           = if ($Det.ValueName) { [string]$Det.ValueName } else { 'DisplayVersion' }
            $normalized.ExpectedValue       = if ($Det.ExpectedValue) { [string]$Det.ExpectedValue } elseif ($Det.DisplayVersion) { [string]$Det.DisplayVersion } else { '' }
            $normalized.Operator            = if ($Det.Operator) { [string]$Det.Operator } else { 'IsEquals' }
            $normalized.PropertyType        = if ($Det.PropertyType) { [string]$Det.PropertyType } else { 'String' }
            $normalized.Is64Bit             = if ($null -ne $Det.Is64Bit) { [bool]$Det.Is64Bit } else { $true }
            $normalized.Hive                = ConvertTo-DTHiveName -Value $rawHive
        }
        'RegistryKey' {
            $rawPath = if ($Det.RegistryKeyRelative) { [string]$Det.RegistryKeyRelative } elseif ($Det.KeyPath) { [string]$Det.KeyPath } else { '' }
            $rawHive = if ($Det.Hive) { [string]$Det.Hive } else { '' }
            if (-not $rawHive) { $rawHive = Get-DTHiveFromPath -Path $rawPath }
            $normalized.RegistryKeyRelative = ConvertTo-DTRegistryRelativePath -Path $rawPath
            $normalized.Is64Bit             = if ($null -ne $Det.Is64Bit) { [bool]$Det.Is64Bit } else { $true }
            $normalized.Hive                = ConvertTo-DTHiveName -Value $rawHive
        }
        'File' {
            $normalized.FilePath      = if ($Det.FilePath) { [string]$Det.FilePath } else { '' }
            $normalized.FileName      = if ($Det.FileName) { [string]$Det.FileName } else { '' }
            $normalized.PropertyType  = if ($Det.PropertyType) { [string]$Det.PropertyType } else { 'Existence' }
            $normalized.ExpectedValue = if ($Det.ExpectedValue) { [string]$Det.ExpectedValue } else { '' }
            $normalized.Operator      = if ($Det.Operator) { [string]$Det.Operator } else { 'GreaterEquals' }
            $normalized.Is64Bit       = if ($null -ne $Det.Is64Bit) { [bool]$Det.Is64Bit } else { $true }
        }
        'Script' {
            $normalized.ScriptText     = if ($Det.ScriptText) { [string]$Det.ScriptText } else { '' }
            $normalized.ScriptLanguage = if ($Det.ScriptLanguage) { [string]$Det.ScriptLanguage } else { 'PowerShell' }
        }
    }

    return $normalized
}

# ---------------------------------------------------------------------------
# Export
# ---------------------------------------------------------------------------

function Export-DetectionResultsCsv {
    <#
    .SYNOPSIS
        Exports a DataTable of test results to CSV.
    #>
    param(
        [Parameter(Mandatory)][System.Data.DataTable]$DataTable,
        [Parameter(Mandatory)][string]$OutputPath
    )

    $parentDir = Split-Path -Path $OutputPath -Parent
    if ($parentDir -and -not (Test-Path -LiteralPath $parentDir)) {
        New-Item -ItemType Directory -Path $parentDir -Force | Out-Null
    }

    $rows = @()
    foreach ($row in $DataTable.Rows) {
        $obj = [ordered]@{}
        foreach ($col in $DataTable.Columns) {
            $obj[$col.ColumnName] = $row[$col.ColumnName]
        }
        $rows += [PSCustomObject]$obj
    }

    $rows | Export-Csv -LiteralPath $OutputPath -NoTypeInformation -Encoding UTF8
    Write-Log "Exported CSV to $OutputPath"
}

function Export-DetectionResultsHtml {
    <#
    .SYNOPSIS
        Exports a DataTable of test results to a self-contained HTML report.
    #>
    param(
        [Parameter(Mandatory)][System.Data.DataTable]$DataTable,
        [Parameter(Mandatory)][string]$OutputPath,
        [string]$ReportTitle = 'Detection Test Results'
    )

    $parentDir = Split-Path -Path $OutputPath -Parent
    if ($parentDir -and -not (Test-Path -LiteralPath $parentDir)) {
        New-Item -ItemType Directory -Path $parentDir -Force | Out-Null
    }

        # Brand: status is conveyed via glyph shape, not color.
    # See feedback_no_red_green_in_brand.md.
    $css = @(
        '<style>',
        'body { font-family: "Segoe UI", Arial, sans-serif; margin: 20px; background: #fafafa; color: #1e1e1e; }',
        'h1 { color: #0078D4; margin-bottom: 4px; }',
        '.summary { color: #666; margin-bottom: 12px; font-size: 0.9em; }',
        'table { border-collapse: collapse; width: 100%; margin-top: 12px; }',
        'th { background: #0078D4; color: #fff; padding: 8px 12px; text-align: left; }',
        'td { padding: 6px 12px; border-bottom: 1px solid #e0e0e0; }',
        'tr:nth-child(even) { background: #f5f5f5; }',
        '.result-pass { font-weight: bold; }',
        '.result-fail { font-weight: bold; }',
        '.glyph { font-family: "Segoe UI Symbol", "Segoe UI", sans-serif; margin-right: 4px; }',
        '</style>'
    ) -join "`r`n"

    # HTML-encode every interpolated value so registry paths / script output
    # containing <, >, &, " can't corrupt the report or execute.
    function ConvertTo-HtmlSafe {
        param([AllowEmptyString()][string]$s)
        if ($null -eq $s) { return '' }
        return [System.Net.WebUtility]::HtmlEncode($s)
    }

    $headerRow = ($DataTable.Columns | ForEach-Object { "<th>$(ConvertTo-HtmlSafe $_.ColumnName)</th>" }) -join ''
    $bodyRows = foreach ($row in $DataTable.Rows) {
        $cells = foreach ($col in $DataTable.Columns) {
            $val = [string]$row[$col.ColumnName]
            $safe = ConvertTo-HtmlSafe $val
            if ($col.ColumnName -eq 'Result') {
                $glyph = if ($val -in @('Pass','Detected')) { '&#x2713;' } else { '&#x2717;' }
                $cssClass = if ($val -in @('Pass','Detected')) { 'result-pass' } else { 'result-fail' }
                "<td class='$cssClass'><span class='glyph'>$glyph</span>$safe</td>"
            } else {
                "<td>$safe</td>"
            }
        }
        "<tr>$($cells -join '')</tr>"
    }

    $titleSafe = ConvertTo-HtmlSafe $ReportTitle
    $html = @(
        '<!DOCTYPE html>',
        '<html><head><meta charset="utf-8"><title>' + $titleSafe + '</title>',
        $css,
        '</head><body>',
        "<h1>$titleSafe</h1>",
        "<div class='summary'>Generated: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') | Tests: $($DataTable.Rows.Count)</div>",
        "<table><thead><tr>$headerRow</tr></thead>",
        "<tbody>$($bodyRows -join "`r`n")</tbody></table>",
        '</body></html>'
    ) -join "`r`n"

    Set-Content -LiteralPath $OutputPath -Value $html -Encoding UTF8
    Write-Log "Exported HTML to $OutputPath"
}
