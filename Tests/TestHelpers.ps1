# Shared helpers for the Detection Method Tester tests. Dot-source from BeforeAll.

$script:AppRoot = Split-Path -Parent $PSScriptRoot

function Import-AppModule {
    Import-Module (Join-Path $script:AppRoot 'Module\DetectionTesterCommon.psd1') -Force -DisableNameChecking
}

# Throwaway registry keys. Every key lives under HKCU\Software\ZZTEST-DT-<guid>, so a crashed run
# leaves one recognisable key and never touches a real value. Remove-TestRegistryKey deletes the
# key in both registry views.
function New-TestRegistryKey {
    param(
        [hashtable]$Value = @{},
        # Parent key under HKCU. The ARP tests use the Uninstall key so that the enumeration finds the entry.
        [string]$Parent = 'Software'
    )
    $relative = $Parent + '\ZZTEST-DT-' + [guid]::NewGuid().ToString('N')
    $base = [Microsoft.Win32.RegistryKey]::OpenBaseKey([Microsoft.Win32.RegistryHive]::CurrentUser, [Microsoft.Win32.RegistryView]::Registry64)
    try {
        $key = $base.CreateSubKey($relative)
        try {
            foreach ($name in $Value.Keys) {
                $entry = $Value[$name]
                if ($entry -is [hashtable]) { $key.SetValue($name, $entry.Data, $entry.Kind) } else { $key.SetValue($name, $entry) }
            }
        }
        finally { $key.Close() }
    }
    finally { $base.Close() }
    return $relative
}

function Remove-TestRegistryKey {
    param([Parameter(Mandatory)][string]$Relative)
    if ($Relative -notlike 'Software\ZZTEST-DT-*' -and $Relative -notlike 'Software\Microsoft\Windows\CurrentVersion\Uninstall\ZZTEST-DT-*') { throw "Refusing to delete '$Relative': not a test key." }
    foreach ($view in [Microsoft.Win32.RegistryView]::Registry64, [Microsoft.Win32.RegistryView]::Registry32) {
        $base = [Microsoft.Win32.RegistryKey]::OpenBaseKey([Microsoft.Win32.RegistryHive]::CurrentUser, $view)
        try { $base.DeleteSubKeyTree($Relative, $false) }
        finally { $base.Close() }
    }
}

# Result table with the columns that the shell passes to the export functions.
function New-ResultTable {
    param([string[]]$Column = @('Time', 'Type', 'Mode', 'Target', 'Expected', 'Found', 'Result'), [object[][]]$Row = @())
    $t = New-Object System.Data.DataTable
    foreach ($c in $Column) { [void]$t.Columns.Add($c, [string]) }
    foreach ($r in $Row) { [void]$t.Rows.Add($r) }
    # The comma keeps PowerShell from unrolling the table into its rows.
    return , $t
}

# The parameter names of one command call: the ones written as -Name, and the keys of a splatted
# hashtable (@params) that the same function or script block builds with a literal or with
# $params['Key'] / $params.Key assignments.
function Get-CommandParameterName {
    param([Parameter(Mandatory)][System.Management.Automation.Language.CommandAst]$Command)
    $names = [System.Collections.Generic.List[string]]::new()
    foreach ($e in $Command.CommandElements) {
        if ($e -is [System.Management.Automation.Language.CommandParameterAst]) { $names.Add($e.ParameterName) }
        elseif ($e -is [System.Management.Automation.Language.VariableExpressionAst] -and $e.Splatted) {
            $var = $e.VariablePath.UserPath
            # Walk outwards from the call to the nearest block (statement block, script block or function) that
            # assigns the variable, so that a variable reused in several switch branches is read per branch.
            $scope = $Command.Parent
            $assigns = @()
            while ($scope) {
                if ($scope -is [System.Management.Automation.Language.StatementBlockAst] -or $scope -is [System.Management.Automation.Language.ScriptBlockAst] -or $scope -is [System.Management.Automation.Language.FunctionDefinitionAst]) {
                    $assigns = @($scope.FindAll({
                                param($n)
                                $n -is [System.Management.Automation.Language.AssignmentStatementAst] -and (
                                    ($n.Left -is [System.Management.Automation.Language.VariableExpressionAst] -and $n.Left.VariablePath.UserPath -eq $var) -or
                                    ($n.Left -is [System.Management.Automation.Language.IndexExpressionAst] -and $n.Left.Target.Extent.Text -eq ('$' + $var)) -or
                                    ($n.Left -is [System.Management.Automation.Language.MemberExpressionAst] -and $n.Left.Expression.Extent.Text -eq ('$' + $var)))
                            }, $true))
                    if ($assigns.Count -gt 0) { break }
                }
                $scope = $scope.Parent
            }
            if ($assigns.Count -eq 0) { continue }
            # A hashtable literal that an outer block assigns before the branch also counts (its keys stay in the splat).
            $outer = $scope.Parent
            while ($outer) {
                if ($outer -is [System.Management.Automation.Language.StatementBlockAst] -or $outer -is [System.Management.Automation.Language.NamedBlockAst]) {
                    $assigns += @($outer.Statements | Where-Object {
                            $_ -is [System.Management.Automation.Language.AssignmentStatementAst] -and $_.Left -is [System.Management.Automation.Language.VariableExpressionAst] -and
                            $_.Left.VariablePath.UserPath -eq $var -and $_.Right.Expression -is [System.Management.Automation.Language.HashtableAst] })
                }
                $outer = $outer.Parent
            }
            foreach ($a in $assigns) {
                $left = $a.Left
                if ($left -is [System.Management.Automation.Language.VariableExpressionAst] -and $left.VariablePath.UserPath -eq $var -and $a.Right.Expression -is [System.Management.Automation.Language.HashtableAst]) {
                    foreach ($kv in $a.Right.Expression.KeyValuePairs) { if ($kv.Item1 -is [System.Management.Automation.Language.StringConstantExpressionAst]) { $names.Add($kv.Item1.Value) } }
                }
                elseif ($left -is [System.Management.Automation.Language.IndexExpressionAst] -and $left.Target -is [System.Management.Automation.Language.VariableExpressionAst] -and $left.Target.VariablePath.UserPath -eq $var -and $left.Index -is [System.Management.Automation.Language.StringConstantExpressionAst]) {
                    $names.Add($left.Index.Value)
                }
                elseif ($left -is [System.Management.Automation.Language.MemberExpressionAst] -and $left.Expression -is [System.Management.Automation.Language.VariableExpressionAst] -and $left.Expression.VariablePath.UserPath -eq $var -and $left.Member -is [System.Management.Automation.Language.StringConstantExpressionAst]) {
                    $names.Add($left.Member.Value)
                }
            }
        }
    }
    return @($names | Select-Object -Unique)
}

# ---- The start script ----------------------------------------------------------------------------------------
# start-detectiontester.ps1 opens the window when it runs, so it cannot be dot-sourced. The tests cut a function,
# or a script-scope script block, out of its syntax tree and define it in the test file.

function Get-ShellAst {
    $path = Join-Path $script:AppRoot 'start-detectiontester.ps1'
    return [System.Management.Automation.Language.Parser]::ParseFile($path, [ref]$null, [ref]$null)
}

# Text of a function definition, with $PSScriptRoot replaced by the repository folder (a function cut out of the
# script has no $PSScriptRoot of its own).
function Get-ShellFunctionText {
    param([Parameter(Mandatory)][string]$Name)
    $def = (Get-ShellAst).Find({ param($n) $n -is [System.Management.Automation.Language.FunctionDefinitionAst] -and $n.Name -eq $Name }, $true)
    if (-not $def) { throw "function $Name not found in the start script" }
    return $def.Extent.Text.Replace('$PSScriptRoot', "'" + ($script:AppRoot -replace "'", "''") + "'")
}

# Body text of `$script:<Name> = { ... }`, as the text between the braces.
function Get-ShellBlockText {
    param([Parameter(Mandatory)][string]$Name)
    $assign = (Get-ShellAst).Find({
            param($n)
            $n -is [System.Management.Automation.Language.AssignmentStatementAst] -and
            $n.Left -is [System.Management.Automation.Language.VariableExpressionAst] -and
            $n.Left.VariablePath.UserPath -eq "script:$Name" -and
            $n.Right.Expression -is [System.Management.Automation.Language.ScriptBlockExpressionAst]
        }, $true)
    if (-not $assign) { throw "script block `$script:$Name not found in the start script" }
    $text = $assign.Right.Expression.Extent.Text
    return $text.Substring(1, $text.Length - 2).Replace('$PSScriptRoot', "'" + ($script:AppRoot -replace "'", "''") + "'")
}

# Names that the start script looks up with <Variable>.FindName('<name>').
function Get-ShellFindNameArgument {
    param([Parameter(Mandatory)][string]$Variable)
    @((Get-ShellAst).FindAll({
                param($n)
                $n -is [System.Management.Automation.Language.InvokeMemberExpressionAst] -and $n.Member.Value -eq 'FindName' -and
                $n.Expression.Extent.Text -eq $Variable -and $n.Arguments.Count -eq 1 -and $n.Arguments[0] -is [System.Management.Automation.Language.StringConstantExpressionAst]
            }, $true) | ForEach-Object { $_.Arguments[0].Value } | Sort-Object -Unique)
}

# Control names that a module loader (`$script:LoadDetectionTester`, `$script:LoadInstalledApps`) takes from its XAML.
function Get-ShellLoaderName {
    param([Parameter(Mandatory)][string]$Loader)
    $assign = (Get-ShellAst).Find({ param($n) $n -is [System.Management.Automation.Language.AssignmentStatementAst] -and $n.Left.Extent.Text -eq ('$script:' + $Loader) }, $true)
    $list = $assign.Right.Find({ param($n) $n -is [System.Management.Automation.Language.ArrayLiteralAst] }, $true)
    @($list.Elements | ForEach-Object { $_.Value })
}

# The parts of the start script that the handler tests run inside a real window (Window.Tests.ps1): the
# handler functions and script blocks. Plain text, so that it can cross to
# the STA runspace.
function Get-ShellRuntimePart {
    param([string[]]$FunctionName = @('Get-DTComboValue', 'Import-DTXaml', 'ConvertTo-DTHistoryDataTable', 'Format-DTAppDetails'))
    $ast = Get-ShellAst
    $root = "'" + ($script:AppRoot -replace "'", "''") + "'"
    $functions = @{}
    foreach ($name in $FunctionName) {
        $def = $ast.Find({ param($n) $n -is [System.Management.Automation.Language.FunctionDefinitionAst] -and $n.Name -eq $name }, $true)
        if (-not $def) { throw "function $name not found in the start script" }
        $functions[$name] = $def.Body.Extent.Text.Substring(1, $def.Body.Extent.Text.Length - 2).Replace('$PSScriptRoot', $root)
    }
    $blocks = @{}
    foreach ($name in 'DTOnTypeChanged', 'DTGatherClause', 'DTValidateClause', 'DTPostTestResult', 'DTRunTest', 'DTRepopulateFromHistory', 'DTExportHistoryCsv', 'DTExportHistoryHtml',
        'DTClearHistory', 'IARefresh', 'IAApplyFilter', 'IAOnAppSelected', 'IAUseForDetection', 'IACopyDetails', 'LoadDetectionTester', 'LoadInstalledApps', 'ShowModule',
        'ApplyTheme') {
        $blocks[$name] = Get-ShellBlockText -Name $name
    }
    # The shell has no palette table to copy: the inactive title bar brush comes from MainWindow.xaml.
    $assignments = @()
    return @{ Functions = $functions; Blocks = $blocks; Assignments = $assignments }
}

# The first 64-bit machine-wide ARP entry that the 32-bit registry view does not show, or $null. Detection of such a key
# must differ between the two views, whatever else is installed.
function Get-ViewExclusiveX64App {
    foreach ($a in (Get-InstalledApplications | Where-Object { $_.Architecture -eq 'x64' -and $_.Scope -eq 'Machine' })) {
        $base = [Microsoft.Win32.RegistryKey]::OpenBaseKey([Microsoft.Win32.RegistryHive]::LocalMachine, [Microsoft.Win32.RegistryView]::Registry32)
        try { $sub = $base.OpenSubKey($a.RegistryKeyRelative); $seen = ($null -ne $sub); if ($sub) { $sub.Close() } }
        finally { $base.Close() }
        if (-not $seen) { return $a }
    }
    return $null
}