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
    .OUTPUTS
        [bool] True if comparison succeeds.
    #>
    param(
        [AllowEmptyString()][AllowNull()][string]$Actual,
        [AllowEmptyString()][AllowNull()][string]$Expected,
        [string]$Operator = 'IsEquals'
    )

    if ($null -eq $Actual) { return $false }

    switch ($Operator) {
        'IsEquals' {
            return ($Actual -eq $Expected)
        }
        'GreaterEquals' {
            try {
                $vActual   = [System.Version]$Actual
                $vExpected = [System.Version]$Expected
                return ($vActual -ge $vExpected)
            } catch {
                Write-Log "Version parse failed: Actual='$Actual', Expected='$Expected' - $_" -Level WARN -Quiet
                return ($Actual -eq $Expected)
            }
        }
        'GreaterThan' {
            try {
                return ([System.Version]$Actual -gt [System.Version]$Expected)
            } catch {
                return $false
            }
        }
        'LessEquals' {
            try {
                return ([System.Version]$Actual -le [System.Version]$Expected)
            } catch {
                return $false
            }
        }
        'LessThan' {
            try {
                return ([System.Version]$Actual -lt [System.Version]$Expected)
            } catch {
                return $false
            }
        }
        default {
            Write-Log "Unknown operator: $Operator" -Level WARN -Quiet
            return $false
        }
    }
}

# ---------------------------------------------------------------------------
# Detection tests
# ---------------------------------------------------------------------------

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
        [bool]$Is64Bit = $true
    )

    $result = @{
        Type          = 'RegistryKeyValue'
        Target        = $RegistryKeyRelative
        Detected      = $false
        KeyExists     = $false
        ValueFound    = $false
        ActualValue   = $null
        ExpectedValue = $ExpectedValue
        Operator      = $Operator
        Details       = ''
    }

    try {
        $view = if ($Is64Bit) {
            [Microsoft.Win32.RegistryView]::Registry64
        } else {
            [Microsoft.Win32.RegistryView]::Registry32
        }

        $hklm = [Microsoft.Win32.RegistryKey]::OpenBaseKey([Microsoft.Win32.RegistryHive]::LocalMachine, $view)
        $subKey = $hklm.OpenSubKey($RegistryKeyRelative)

        if ($null -eq $subKey) {
            $result.Details = "Key not found: HKLM\$RegistryKeyRelative"
            Write-Log "RegistryKeyValue: Key not found - HKLM\$RegistryKeyRelative" -Quiet
            return $result
        }

        $result.KeyExists = $true
        $rawValue = $subKey.GetValue($ValueName, $null)
        $subKey.Close()
        $hklm.Close()

        if ($null -eq $rawValue) {
            $result.Details = "Value '$ValueName' not found in key"
            Write-Log "RegistryKeyValue: Value '$ValueName' not found" -Quiet
            return $result
        }

        $result.ValueFound = $true
        $result.ActualValue = [string]$rawValue

        $match = Compare-DetectionValue -Actual ([string]$rawValue) -Expected $ExpectedValue -Operator $Operator
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
        [bool]$Is64Bit = $true
    )

    $result = @{
        Type          = 'RegistryKey'
        Target        = $RegistryKeyRelative
        Detected      = $false
        KeyExists     = $false
        ActualValue   = $null
        ExpectedValue = '(key exists)'
        Operator      = 'Existence'
        Details       = ''
    }

    try {
        $view = if ($Is64Bit) {
            [Microsoft.Win32.RegistryView]::Registry64
        } else {
            [Microsoft.Win32.RegistryView]::Registry32
        }

        $hklm = [Microsoft.Win32.RegistryKey]::OpenBaseKey([Microsoft.Win32.RegistryHive]::LocalMachine, $view)
        $subKey = $hklm.OpenSubKey($RegistryKeyRelative)

        if ($null -ne $subKey) {
            $result.Detected = $true
            $result.KeyExists = $true
            $result.Details = "Key exists: HKLM\$RegistryKeyRelative"
            $subKey.Close()
        } else {
            $result.Details = "Key not found: HKLM\$RegistryKeyRelative"
        }
        $hklm.Close()

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

            $match = Compare-DetectionValue -Actual $fileVersion -Expected $ExpectedValue -Operator $Operator
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
    .OUTPUTS
        [hashtable] with keys: Detected, ScriptOutput, Details
    #>
    param(
        [Parameter(Mandatory)][string]$ScriptText,
        [string]$ScriptLanguage = 'PowerShell'
    )

    $result = @{
        Type          = 'Script'
        Target        = '(script)'
        Detected      = $false
        ActualValue   = $null
        ExpectedValue = '(non-empty output)'
        Operator      = 'Script'
        ScriptOutput  = $null
        Details       = ''
    }

    if ($ScriptLanguage -ne 'PowerShell') {
        $result.Details = "Unsupported script language: $ScriptLanguage (only PowerShell supported)"
        Write-Log "Script: $($result.Details)" -Level WARN -Quiet
        return $result
    }

    try {
        $sb = [scriptblock]::Create($ScriptText)
        $output = & $sb 2>&1

        $stdout = @($output | Where-Object { $_ -isnot [System.Management.Automation.ErrorRecord] }) -join "`r`n"
        $stderr = @($output | Where-Object { $_ -is [System.Management.Automation.ErrorRecord] }) -join "`r`n"

        $result.ScriptOutput = $stdout
        $result.ActualValue = $stdout

        if (-not [string]::IsNullOrWhiteSpace($stdout)) {
            $result.Detected = $true
            $result.Details = "Script produced output (DETECTED): $($stdout.Substring(0, [Math]::Min(200, $stdout.Length)))"
        } else {
            $result.Details = "Script produced no output (NOT DETECTED)"
            if (-not [string]::IsNullOrWhiteSpace($stderr)) {
                $result.Details += " | Errors: $($stderr.Substring(0, [Math]::Min(200, $stderr.Length)))"
            }
        }

        Write-Log "Script: $($result.Details)" -Quiet
    } catch {
        $result.Details = "Script execution error: $_"
        Write-Log "Script error: $_" -Level ERROR -Quiet
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
                }
                Test-RegistryKeyValueDetection @p
            }
            'RegistryKey' {
                $p = @{
                    RegistryKeyRelative = $clause.RegistryKeyRelative
                    Is64Bit             = if ($null -ne $clause.Is64Bit) { [bool]$clause.Is64Bit } else { $true }
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
    $arpPaths = @(
        @{ Path = 'SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall'; Arch = 'x64' }
        @{ Path = 'SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall'; Arch = 'x86' }
    )

    foreach ($arp in $arpPaths) {
        try {
            $hklm = [Microsoft.Win32.RegistryKey]::OpenBaseKey(
                [Microsoft.Win32.RegistryHive]::LocalMachine,
                [Microsoft.Win32.RegistryView]::Registry64
            )
            $parentKey = $hklm.OpenSubKey($arp.Path)
            if ($null -eq $parentKey) { continue }

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
                        RegistryKey          = "$($arp.Path)\$subKeyName"
                        UninstallString      = [string]$subKey.GetValue('UninstallString', '')
                        QuietUninstallString = [string]$subKey.GetValue('QuietUninstallString', '')
                        InstallLocation      = [string]$subKey.GetValue('InstallLocation', '')
                        InstallDate          = [string]$subKey.GetValue('InstallDate', '')
                    }

                    $subKey.Close()
                } catch {
                    # Access denied on individual keys is normal
                }
            }

            $parentKey.Close()
            $hklm.Close()
        } catch {
            Write-Log "ARP enumeration error ($($arp.Arch)): $_" -Level WARN -Quiet
        }
    }

    Write-Log "Enumerated $($apps.Count) installed applications" -Quiet
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
            $normalized.RegistryKeyRelative = if ($Det.RegistryKeyRelative) { [string]$Det.RegistryKeyRelative } elseif ($Det.KeyPath) { [string]$Det.KeyPath } else { '' }
            $normalized.ValueName           = if ($Det.ValueName) { [string]$Det.ValueName } else { 'DisplayVersion' }
            $normalized.ExpectedValue       = if ($Det.ExpectedValue) { [string]$Det.ExpectedValue } elseif ($Det.DisplayVersion) { [string]$Det.DisplayVersion } else { '' }
            $normalized.Operator            = if ($Det.Operator) { [string]$Det.Operator } else { 'IsEquals' }
            $normalized.PropertyType        = if ($Det.PropertyType) { [string]$Det.PropertyType } else { 'String' }
            $normalized.Is64Bit             = if ($null -ne $Det.Is64Bit) { [bool]$Det.Is64Bit } else { $true }
        }
        'RegistryKey' {
            $normalized.RegistryKeyRelative = if ($Det.RegistryKeyRelative) { [string]$Det.RegistryKeyRelative } elseif ($Det.KeyPath) { [string]$Det.KeyPath } else { '' }
            $normalized.Is64Bit             = if ($null -ne $Det.Is64Bit) { [bool]$Det.Is64Bit } else { $true }
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

    $css = @(
        '<style>',
        'body { font-family: "Segoe UI", Arial, sans-serif; margin: 20px; background: #fafafa; }',
        'h1 { color: #0078D4; margin-bottom: 4px; }',
        '.summary { color: #666; margin-bottom: 12px; font-size: 0.9em; }',
        'table { border-collapse: collapse; width: 100%; margin-top: 12px; }',
        'th { background: #0078D4; color: #fff; padding: 8px 12px; text-align: left; }',
        'td { padding: 6px 12px; border-bottom: 1px solid #e0e0e0; }',
        'tr:nth-child(even) { background: #f5f5f5; }',
        '.detected { color: #228B22; font-weight: bold; }',
        '.not-detected { color: #B40000; font-weight: bold; }',
        '</style>'
    ) -join "`r`n"

    $headerRow = ($DataTable.Columns | ForEach-Object { "<th>$($_.ColumnName)</th>" }) -join ''
    $bodyRows = foreach ($row in $DataTable.Rows) {
        $cells = foreach ($col in $DataTable.Columns) {
            $val = [string]$row[$col.ColumnName]
            if ($col.ColumnName -eq 'Result') {
                $cssClass = if ($val -eq 'Detected') { 'detected' } else { 'not-detected' }
                "<td class='$cssClass'>$val</td>"
            } else {
                "<td>$val</td>"
            }
        }
        "<tr>$($cells -join '')</tr>"
    }

    $html = @(
        '<!DOCTYPE html>',
        '<html><head><meta charset="utf-8"><title>' + $ReportTitle + '</title>',
        $css,
        '</head><body>',
        "<h1>$ReportTitle</h1>",
        "<div class='summary'>Generated: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') | Tests: $($DataTable.Rows.Count)</div>",
        "<table><thead><tr>$headerRow</tr></thead>",
        "<tbody>$($bodyRows -join "`r`n")</tbody></table>",
        '</body></html>'
    ) -join "`r`n"

    Set-Content -LiteralPath $OutputPath -Value $html -Encoding UTF8
    Write-Log "Exported HTML to $OutputPath"
}
