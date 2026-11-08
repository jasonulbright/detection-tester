#Requires -Modules Pester
# Logic of start-detectiontester.ps1. The script opens its window when it runs, so it cannot be dot-sourced.
# Each function and each script-scope script block is cut out of its syntax tree and defined here, with plain
# objects in place of the window controls. The real controls are used in Window.Tests.ps1.

BeforeAll {
    . "$PSScriptRoot\TestHelpers.ps1"
    Import-AppModule
    Add-Type -AssemblyName PresentationFramework, PresentationCore, WindowsBase

    foreach ($name in 'Get-DTComboValue', 'ConvertTo-DTHistoryDataTable', 'Get-DTPrefs', 'Save-DTPrefs', 'Restore-DTWindowState', 'Save-DTWindowState') {
        . ([scriptblock]::Create((Get-ShellFunctionText -Name $name)))
    }
    # The blocks call each other through $script:<Name>, so each one is also a script variable.
    $script:DTOnTypeChanged = [scriptblock]::Create((Get-ShellBlockText -Name 'DTOnTypeChanged'))
    $script:DTGatherClause = [scriptblock]::Create((Get-ShellBlockText -Name 'DTGatherClause'))
    $script:DTValidateClause = [scriptblock]::Create((Get-ShellBlockText -Name 'DTValidateClause'))
    $script:DTPostTestResult = [scriptblock]::Create((Get-ShellBlockText -Name 'DTPostTestResult'))
    $script:DTRunTest = [scriptblock]::Create((Get-ShellBlockText -Name 'DTRunTest'))
    $script:DTImportManifest = [scriptblock]::Create((Get-ShellBlockText -Name 'DTImportManifest'))
    $script:DTRepopulateFromHistory = [scriptblock]::Create((Get-ShellBlockText -Name 'DTRepopulateFromHistory'))
    $script:DTExportHistoryCsv = [scriptblock]::Create((Get-ShellBlockText -Name 'DTExportHistoryCsv'))
    $script:DTExportHistoryHtml = [scriptblock]::Create((Get-ShellBlockText -Name 'DTExportHistoryHtml'))
    $script:DTClearHistory = [scriptblock]::Create((Get-ShellBlockText -Name 'DTClearHistory'))
    $script:IARefresh = [scriptblock]::Create((Get-ShellBlockText -Name 'IARefresh'))
    $script:IAApplyFilter = [scriptblock]::Create((Get-ShellBlockText -Name 'IAApplyFilter'))
    $script:IAOnAppSelected = [scriptblock]::Create((Get-ShellBlockText -Name 'IAOnAppSelected'))
    $script:IAUseForDetection = [scriptblock]::Create((Get-ShellBlockText -Name 'IAUseForDetection'))
    $script:ShowModule = [scriptblock]::Create((Get-ShellBlockText -Name 'ShowModule'))
    $script:OnClosing = [scriptblock]::Create((Get-ShellBlockText -Name 'OnClosing'))
    $script:Blocks = @{ DTExportHistoryCsv = $script:DTExportHistoryCsv; DTExportHistoryHtml = $script:DTExportHistoryHtml }

    function script:New-Combo {
        param([string[]]$Item, [int]$Index = 0)
        $items = New-Object System.Collections.ArrayList
        foreach ($i in $Item) { $c = New-Object System.Windows.Controls.ComboBoxItem; $c.Content = $i; [void]$items.Add($c) }
        $o = [pscustomobject]@{ Items = $items; SelectedIndex = $Index }
        Add-Member -InputObject $o -MemberType ScriptProperty -Name SelectedItem -Value { if ($this.SelectedIndex -ge 0 -and $this.SelectedIndex -lt $this.Items.Count) { $this.Items[$this.SelectedIndex] } else { $null } }
        return $o
    }
    function script:New-Context {
        $operators = 'IsEquals', 'GreaterEquals', 'GreaterThan', 'LessEquals', 'LessThan'
        $ctx = @{
            cboType = New-Combo -Item 'RegistryKeyValue', 'RegistryKey', 'File', 'Script', 'Compound (import-only)'
            chkNegativeMode = [pscustomobject]@{ IsChecked = $false }
            btnTest = [pscustomobject]@{ IsEnabled = $true }
            txtResultBadge = [pscustomobject]@{ Text = '' }
            txtResultDetails = [pscustomobject]@{ Text = '' }
            btnRepopulate = [pscustomobject]@{ Visibility = 'Collapsed' }
            dgHistory = [pscustomobject]@{ SelectedItem = $null }
            rkvHiveHKLM = [pscustomobject]@{ IsChecked = $true }; rkvHiveHKCU = [pscustomobject]@{ IsChecked = $false }
            rkvView64 = [pscustomobject]@{ IsChecked = $true }; rkvView32 = [pscustomobject]@{ IsChecked = $false }
            rkvKeyPath = [pscustomobject]@{ Text = '' }; rkvValueName = [pscustomobject]@{ Text = '' }
            rkvOperator = New-Combo -Item $operators; rkvExpected = [pscustomobject]@{ Text = '' }
            rkHiveHKLM = [pscustomobject]@{ IsChecked = $true }; rkHiveHKCU = [pscustomobject]@{ IsChecked = $false }
            rkView64 = [pscustomobject]@{ IsChecked = $true }; rkView32 = [pscustomobject]@{ IsChecked = $false }
            rkKeyPath = [pscustomobject]@{ Text = '' }
            fileView64 = [pscustomobject]@{ IsChecked = $true }; fileView32 = [pscustomobject]@{ IsChecked = $false }
            fileFilePath = [pscustomobject]@{ Text = '' }; fileFileName = [pscustomobject]@{ Text = '' }
            filePropType = New-Combo -Item 'Existence', 'Version'; fileOperator = New-Combo -Item 'GreaterEquals', 'GreaterThan', 'IsEquals', 'LessEquals', 'LessThan'
            fileExpected = [pscustomobject]@{ Text = '' }
            scriptText = [pscustomobject]@{ Text = '' }
            compoundConnector = [pscustomobject]@{ Text = '' }
            compoundClauses = [pscustomobject]@{ Items = (New-Object System.Collections.ArrayList) }
            panelRegistryKeyValue = [pscustomobject]@{ Visibility = 'Visible' }; panelRegistryKey = [pscustomobject]@{ Visibility = 'Collapsed' }
            panelFile = [pscustomobject]@{ Visibility = 'Collapsed' }; panelScript = [pscustomobject]@{ Visibility = 'Collapsed' }
            panelCompound = [pscustomobject]@{ Visibility = 'Collapsed' }
        }
        return $ctx
    }
    $global:Fx = @{ Calls = (New-Object System.Collections.ArrayList) }
}

AfterAll { Remove-Variable -Name Fx -Scope Global -ErrorAction SilentlyContinue }

Describe 'Get-DTComboValue' {
    It 'gives an empty string when nothing is selected' {
        Get-DTComboValue -Combo ([pscustomobject]@{ SelectedItem = $null }) | Should -Be ''
    }
    It 'gives the content of a combo box item' {
        $item = New-Object System.Windows.Controls.ComboBoxItem; $item.Content = 'GreaterEquals'
        Get-DTComboValue -Combo ([pscustomobject]@{ SelectedItem = $item }) | Should -Be 'GreaterEquals'
    }
    It 'gives an empty string for an item without content' {
        $item = New-Object System.Windows.Controls.ComboBoxItem
        Get-DTComboValue -Combo ([pscustomobject]@{ SelectedItem = $item }) | Should -Be ''
    }
    It 'gives a plain string as it is' {
        Get-DTComboValue -Combo ([pscustomobject]@{ SelectedItem = 'Machine (HKLM)' }) | Should -Be 'Machine (HKLM)'
    }
}

Describe 'ConvertTo-DTHistoryDataTable' {
    It 'returns one DataTable for <Count> rows' -ForEach @(@{ Count = 0 }, @{ Count = 1 }, @{ Count = 3 }) {
        $item = [pscustomobject]@{ Time = '12:00:00'; Type = 'File'; Mode = 'Positive'; Target = 'a'; Expected = 'b'; Found = 'c'; Result = 'Pass' }
        $table = ConvertTo-DTHistoryDataTable -Items @(for ($i = 0; $i -lt $Count; $i++) { $item })
        $table.GetType().FullName | Should -Be 'System.Data.DataTable'
        $table.Rows.Count | Should -Be $Count
        ($table.Columns.ColumnName -join ',') | Should -Be 'Time,Type,Mode,Target,Expected,Found,Result'
    }

    It 'copies the values of a row' {
        $item = [pscustomobject]@{ Time = '12:00:01'; Type = 'Script'; Mode = 'Negative'; Target = '(script)'; Expected = '(non-empty output)'; Found = 'x'; Result = 'Fail'; Clause = @{ a = 1 }; Negative = $true }
        $table = ConvertTo-DTHistoryDataTable -Items @($item)
        $table.Rows[0].Mode | Should -Be 'Negative'
        $table.Rows[0].Found | Should -Be 'x'
        $table.Columns.Contains('Clause') | Should -BeFalse
    }

    It 'turns a missing value into an empty cell' {
        $table = ConvertTo-DTHistoryDataTable -Items @([pscustomobject]@{ Time = 't'; Result = $null })
        $table.Rows[0].Target | Should -Be ''
        $table.Rows[0].Result | Should -Be ''
    }

    It 'accepts the observable collection that the window uses' {
        $items = New-Object 'System.Collections.ObjectModel.ObservableCollection[psobject]'
        $items.Add([pscustomobject]@{ Time = 't'; Type = 'File'; Mode = 'Positive'; Target = 'a'; Expected = 'b'; Found = 'c'; Result = 'Pass' })
        (ConvertTo-DTHistoryDataTable -Items $items).Rows.Count | Should -Be 1
    }
}

Describe 'Preferences' {
    BeforeEach {
        $script:prefsPath = Join-Path $TestDrive ('prefs-' + [guid]::NewGuid().ToString('N') + '.json')
    }

    It 'gives the defaults when the file does not exist' {
        $p = Get-DTPrefs
        $p.DarkMode | Should -BeTrue
        $p.ActiveModule | Should -Be 'DetectionTester'
    }

    It 'saves and reads the values back' {
        Save-DTPrefs -Prefs @{ DarkMode = $false; ActiveModule = 'InstalledApps' }
        Test-Path -LiteralPath $script:prefsPath | Should -BeTrue
        $p = Get-DTPrefs
        $p.DarkMode | Should -BeFalse
        $p.ActiveModule | Should -Be 'InstalledApps'
    }

    It 'gives the defaults when the file holds text that is not JSON' {
        Set-Content -LiteralPath $script:prefsPath -Value '{ broken'
        $p = Get-DTPrefs
        $p.DarkMode | Should -BeTrue
        $p.ActiveModule | Should -Be 'DetectionTester'
    }

    It 'adds the default for a value that the file does not have' {
        Set-Content -LiteralPath $script:prefsPath -Value '{"DarkMode": false}'
        $p = Get-DTPrefs
        $p.DarkMode | Should -BeFalse
        $p.ActiveModule | Should -Be 'DetectionTester'
    }
}

Describe 'Window state' {
    BeforeEach {
        Mock Write-Log {}
        $script:wsPath = Join-Path $TestDrive ('ws-' + [guid]::NewGuid().ToString('N') + '.json')
        $script:window = [pscustomobject]@{
            Width = 1100; Height = 750; Left = 10; Top = 10; ActualWidth = 1100; ActualHeight = 750
            WindowState = [System.Windows.WindowState]::Normal; WindowStartupLocation = [System.Windows.WindowStartupLocation]::CenterScreen
            RestoreBounds = New-Object System.Windows.Rect(50, 60, 1000, 700)
        }
        $script:vsLeft = [System.Windows.SystemParameters]::VirtualScreenLeft
        $script:vsTop = [System.Windows.SystemParameters]::VirtualScreenTop
    }

    It 'changes nothing when there is no saved state' {
        Restore-DTWindowState
        $script:window.Width | Should -Be 1100
        $script:window.WindowStartupLocation | Should -Be ([System.Windows.WindowStartupLocation]::CenterScreen)
    }

    It 'restores a saved size and position that are on screen' {
        $left = $script:vsLeft + 20; $top = $script:vsTop + 20
        Set-Content -LiteralPath $script:wsPath -Value (@{ Left = $left; Top = $top; Width = 900; Height = 650; Maximized = $false } | ConvertTo-Json)
        Restore-DTWindowState
        $script:window.Width | Should -Be 900
        $script:window.Height | Should -Be 650
        $script:window.Left | Should -Be $left
        $script:window.Top | Should -Be $top
        $script:window.WindowStartupLocation | Should -Be ([System.Windows.WindowStartupLocation]::Manual)
        $script:window.WindowState | Should -Be ([System.Windows.WindowState]::Normal)
    }

    It 'centres the window when the saved position is off screen' {
        Set-Content -LiteralPath $script:wsPath -Value (@{ Left = 90000; Top = 90000; Width = 900; Height = 650; Maximized = $false } | ConvertTo-Json)
        Restore-DTWindowState
        $script:window.WindowStartupLocation | Should -Be ([System.Windows.WindowStartupLocation]::CenterScreen)
        $script:window.Width | Should -Be 900
        $script:window.Left | Should -Be 10
    }

    It 'centres the window when the saved position is the minimized position' {
        Set-Content -LiteralPath $script:wsPath -Value (@{ Left = -32000; Top = -32000; Width = 900; Height = 650; Maximized = $false } | ConvertTo-Json)
        Restore-DTWindowState
        $script:window.WindowStartupLocation | Should -Be ([System.Windows.WindowStartupLocation]::CenterScreen)
    }

    It 'keeps the default size when the saved size is too small: <Width>x<Height>' -ForEach @(
        @{ Width = 600; Height = 700 }, @{ Width = 900; Height = 400 }, @{ Width = 0; Height = 0 }, @{ Width = -5; Height = -5 }
    ) {
        Set-Content -LiteralPath $script:wsPath -Value (@{ Left = $script:vsLeft + 20; Top = $script:vsTop + 20; Width = $Width; Height = $Height; Maximized = $false } | ConvertTo-Json)
        Restore-DTWindowState
        if ($Width -le 600) { $script:window.Width | Should -Be 1100 } else { $script:window.Width | Should -Be $Width }
        if ($Height -le 400) { $script:window.Height | Should -Be 750 } else { $script:window.Height | Should -Be $Height }
    }

    It 'maximizes the window when the saved state says so' {
        Set-Content -LiteralPath $script:wsPath -Value (@{ Left = $script:vsLeft + 20; Top = $script:vsTop + 20; Width = 900; Height = 650; Maximized = $true } | ConvertTo-Json)
        Restore-DTWindowState
        $script:window.WindowState | Should -Be ([System.Windows.WindowState]::Maximized)
    }

    It 'does not throw and logs when the saved state is damaged' {
        Set-Content -LiteralPath $script:wsPath -Value 'not json'
        { Restore-DTWindowState } | Should -Not -Throw
        Should -Invoke Write-Log -Times 1 -Exactly -ParameterFilter { $Message -like 'Restore-DTWindowState:*' }
        $script:window.Width | Should -Be 1100
    }

    It 'saves the size of a normal window' {
        $script:window.Left = 33; $script:window.Top = 44; $script:window.ActualWidth = 1234; $script:window.ActualHeight = 777
        Save-DTWindowState
        $s = Get-Content -LiteralPath $script:wsPath -Raw | ConvertFrom-Json
        $s.Left | Should -Be 33
        $s.Top | Should -Be 44
        $s.Width | Should -Be 1234
        $s.Height | Should -Be 777
        $s.Maximized | Should -BeFalse
    }

    It 'saves the restore bounds of a maximized window' {
        $script:window.WindowState = [System.Windows.WindowState]::Maximized
        $script:window.ActualWidth = 1920
        Save-DTWindowState
        $s = Get-Content -LiteralPath $script:wsPath -Raw | ConvertFrom-Json
        $s.Maximized | Should -BeTrue
        $s.Width | Should -Be 1000
        $s.Height | Should -Be 700
        $s.Left | Should -Be 50
    }

    It 'restores what it saved' {
        $script:window.Left = $script:vsLeft + 15; $script:window.Top = $script:vsTop + 25; $script:window.ActualWidth = 950; $script:window.ActualHeight = 640
        Save-DTWindowState
        $script:window.Width = 1; $script:window.Height = 1
        Restore-DTWindowState
        $script:window.Width | Should -Be 950
        $script:window.Height | Should -Be 640
        $script:window.Left | Should -Be ($script:vsLeft + 15)
    }
}

Describe 'DTOnTypeChanged' {
    BeforeEach { $script:DTContext = New-Context }

    It 'shows only the panel of <Type>' -ForEach @(
        @{ Index = 0; Type = 'RegistryKeyValue'; Panel = 'panelRegistryKeyValue' }
        @{ Index = 1; Type = 'RegistryKey'; Panel = 'panelRegistryKey' }
        @{ Index = 2; Type = 'File'; Panel = 'panelFile' }
        @{ Index = 3; Type = 'Script'; Panel = 'panelScript' }
        @{ Index = 4; Type = 'Compound (import-only)'; Panel = 'panelCompound' }
    ) {
        $script:DTContext.cboType.SelectedIndex = $Index
        & $script:DTOnTypeChanged
        foreach ($p in 'panelRegistryKeyValue', 'panelRegistryKey', 'panelFile', 'panelScript', 'panelCompound') {
            $script:DTContext[$p].Visibility | Should -Be $(if ($p -eq $Panel) { 'Visible' } else { 'Collapsed' })
        }
    }

    It 'hides every panel when no type is selected' {
        $script:DTContext.cboType.SelectedIndex = -1
        & $script:DTOnTypeChanged
        foreach ($p in 'panelRegistryKeyValue', 'panelRegistryKey', 'panelFile', 'panelScript', 'panelCompound') { $script:DTContext[$p].Visibility | Should -Be 'Collapsed' }
    }

    It 'does nothing before the controls exist' {
        $script:DTContext = @{}
        { & $script:DTOnTypeChanged } | Should -Not -Throw
    }
}

Describe 'DTGatherClause' {
    BeforeEach { $script:DTContext = New-Context }

    It 'gathers a registry value clause' {
        $c = $script:DTContext
        $c.cboType.SelectedIndex = 0
        $c.rkvHiveHKCU.IsChecked = $true; $c.rkvHiveHKLM.IsChecked = $false
        $c.rkvView64.IsChecked = $false; $c.rkvView32.IsChecked = $true
        $c.rkvKeyPath.Text = 'Software\V'; $c.rkvValueName.Text = 'Version'; $c.rkvOperator.SelectedIndex = 2; $c.rkvExpected.Text = '1.0'
        $g = & $script:DTGatherClause
        $g.Type | Should -Be 'RegistryKeyValue'
        $g.Hive | Should -Be 'CurrentUser'
        $g.Is64Bit | Should -BeFalse
        $g.RegistryKeyRelative | Should -Be 'Software\V'
        $g.ValueName | Should -Be 'Version'
        $g.Operator | Should -Be 'GreaterThan'
        $g.ExpectedValue | Should -Be '1.0'
        $g.PropertyType | Should -Be 'String'
    }

    It 'gathers a registry key clause, with HKLM and the 64-bit view as the defaults' {
        $c = $script:DTContext
        $c.cboType.SelectedIndex = 1
        $c.rkKeyPath.Text = 'SOFTWARE\V'
        $g = & $script:DTGatherClause
        $g.Type | Should -Be 'RegistryKey'
        $g.Hive | Should -Be 'LocalMachine'
        $g.Is64Bit | Should -BeTrue
        $g.RegistryKeyRelative | Should -Be 'SOFTWARE\V'
    }

    It 'gathers a file clause' {
        $c = $script:DTContext
        $c.cboType.SelectedIndex = 2
        $c.fileFilePath.Text = 'C:\Program Files\V'; $c.fileFileName.Text = 'v.exe'; $c.filePropType.SelectedIndex = 1; $c.fileOperator.SelectedIndex = 4; $c.fileExpected.Text = '2.0'
        $c.fileView64.IsChecked = $false
        $g = & $script:DTGatherClause
        $g.Type | Should -Be 'File'
        $g.FilePath | Should -Be 'C:\Program Files\V'
        $g.FileName | Should -Be 'v.exe'
        $g.PropertyType | Should -Be 'Version'
        $g.Operator | Should -Be 'LessThan'
        $g.ExpectedValue | Should -Be '2.0'
        $g.Is64Bit | Should -BeFalse
    }

    It 'gathers a script clause' {
        $script:DTContext.cboType.SelectedIndex = 3
        $script:DTContext.scriptText.Text = '"found"'
        $g = & $script:DTGatherClause
        $g.Type | Should -Be 'Script'
        $g.ScriptText | Should -Be '"found"'
        $g.ScriptLanguage | Should -Be 'PowerShell'
    }

    It 'gives nothing for a compound clause that has no imported data' {
        $script:DTContext.cboType.SelectedIndex = 4
        & $script:DTGatherClause | Should -BeNullOrEmpty
        $script:DTContext.CompoundClauseData = @()
        & $script:DTGatherClause | Should -BeNullOrEmpty
    }

    It 'gathers a compound clause with the imported data' {
        $script:DTContext.cboType.SelectedIndex = 4
        $script:DTContext.compoundConnector.Text = 'Or'
        $script:DTContext.CompoundClauseData = @(@{ Type = 'RegistryKey' }, @{ Type = 'File' })
        $g = & $script:DTGatherClause
        $g.Type | Should -Be 'Compound'
        $g.Connector | Should -Be 'Or'
        $g.Clauses.Count | Should -Be 2
    }

    It 'gives nothing when no type is selected' {
        $script:DTContext.cboType.SelectedIndex = -1
        & $script:DTGatherClause | Should -BeNullOrEmpty
    }
}

Describe 'DTValidateClause' {
    It 'answers "<Message>" for <Case>' -ForEach @(
        @{ Case = 'no clause';                         Clause = $null;                                                                         Message = 'No clause to evaluate.' }
        @{ Case = 'a registry value without a key';    Clause = @{ Type = 'RegistryKeyValue'; RegistryKeyRelative = '';  ValueName = 'v' };          Message = 'Key path is required.' }
        @{ Case = 'a registry value with a blank key'; Clause = @{ Type = 'RegistryKeyValue'; RegistryKeyRelative = '  '; ValueName = 'v' };         Message = 'Key path is required.' }
        @{ Case = 'a registry value without a name';   Clause = @{ Type = 'RegistryKeyValue'; RegistryKeyRelative = 'k'; ValueName = '' };           Message = 'Value name is required.' }
        @{ Case = 'a registry key without a path';     Clause = @{ Type = 'RegistryKey'; RegistryKeyRelative = '' };                                  Message = 'Key path is required.' }
        @{ Case = 'a file without a folder';           Clause = @{ Type = 'File'; FilePath = '';   FileName = 'a.exe'; PropertyType = 'Existence' }; Message = 'File path is required.' }
        @{ Case = 'a file without a name';             Clause = @{ Type = 'File'; FilePath = 'C:\'; FileName = '';      PropertyType = 'Existence' }; Message = 'File name is required.' }
        @{ Case = 'a file version without a version';  Clause = @{ Type = 'File'; FilePath = 'C:\'; FileName = 'a.exe'; PropertyType = 'Version'; ExpectedValue = ' ' }; Message = 'Expected version is required when Property type = Version.' }
        @{ Case = 'a script without text';             Clause = @{ Type = 'Script'; ScriptText = '   ' };                                            Message = 'Script body is required.' }
        @{ Case = 'a compound without clauses';        Clause = @{ Type = 'Compound'; Clauses = @() };                                               Message = 'Compound clause has no children (import a manifest with Clauses).' }
        @{ Case = 'a compound with null clauses';      Clause = @{ Type = 'Compound'; Clauses = $null };                                             Message = 'Compound clause has no children (import a manifest with Clauses).' }
    ) {
        & $script:DTValidateClause -Clause $Clause | Should -Be $Message
    }

    It 'accepts <Case>' -ForEach @(
        @{ Case = 'a registry value';              Clause = @{ Type = 'RegistryKeyValue'; RegistryKeyRelative = 'k'; ValueName = 'v'; ExpectedValue = '' } }
        @{ Case = 'a registry key';                Clause = @{ Type = 'RegistryKey'; RegistryKeyRelative = 'k' } }
        @{ Case = 'a file by existence';           Clause = @{ Type = 'File'; FilePath = 'C:\'; FileName = 'a'; PropertyType = 'Existence'; ExpectedValue = '' } }
        @{ Case = 'a file by version';             Clause = @{ Type = 'File'; FilePath = 'C:\'; FileName = 'a'; PropertyType = 'Version'; ExpectedValue = '1.0' } }
        @{ Case = 'a script';                      Clause = @{ Type = 'Script'; ScriptText = '"x"' } }
        @{ Case = 'a compound with a clause';      Clause = @{ Type = 'Compound'; Clauses = @(@{ Type = 'RegistryKey' }) } }
    ) {
        & $script:DTValidateClause -Clause $Clause | Should -BeNullOrEmpty
    }
}

Describe 'DTPostTestResult' {
    BeforeEach {
        Mock Write-Log {}
        $script:DTContext = New-Context
        $script:HistoryItems = New-Object 'System.Collections.ObjectModel.ObservableCollection[psobject]'
    }

    It 'shows <Badge> for Detected=<Detected> in <Mode> mode' -ForEach @(
        @{ Detected = $true;  Negative = $false; Badge = 'Pass'; Mode = 'Positive' }
        @{ Detected = $false; Negative = $false; Badge = 'Fail'; Mode = 'Positive' }
        @{ Detected = $false; Negative = $true;  Badge = 'Pass'; Mode = 'Negative' }
        @{ Detected = $true;  Negative = $true;  Badge = 'Fail'; Mode = 'Negative' }
    ) {
        $result = @{ Type = 'RegistryKey'; Detected = $Detected; Details = 'the details'; Target = 'HKLM\X'; ExpectedValue = '(key exists)'; ActualValue = $null }
        & $script:DTPostTestResult -Result $result -Negative $Negative -Clause @{ Type = 'RegistryKey' }
        $script:DTContext.txtResultBadge.Text | Should -Be $Badge
        $script:DTContext.txtResultDetails.Text | Should -Be 'the details'
        $script:HistoryItems.Count | Should -Be 1
        $script:HistoryItems[0].Result | Should -Be $Badge
        $script:HistoryItems[0].Mode | Should -Be $Mode
        $script:HistoryItems[0].Negative | Should -Be $Negative
    }

    It 'puts the newest result first and keeps the clause snapshot' {
        $clause1 = @{ Type = 'File'; FileName = 'one' }
        $clause2 = @{ Type = 'File'; FileName = 'two' }
        & $script:DTPostTestResult -Result @{ Type = 'File'; Detected = $true; Details = ''; Target = 'one'; ExpectedValue = 'e'; ActualValue = 'a' } -Negative $false -Clause $clause1
        & $script:DTPostTestResult -Result @{ Type = 'File'; Detected = $false; Details = ''; Target = 'two'; ExpectedValue = 'e'; ActualValue = $null } -Negative $false -Clause $clause2
        $script:HistoryItems.Count | Should -Be 2
        $script:HistoryItems[0].Target | Should -Be 'two'
        $script:HistoryItems[0].Clause.FileName | Should -Be 'two'
        $script:HistoryItems[1].Clause.FileName | Should -Be 'one'
    }

    It 'fills the columns as text and uses an empty text for a missing actual value' {
        & $script:DTPostTestResult -Result @{ Type = 'File'; Detected = $true; Details = ''; Target = 'T'; ExpectedValue = 5; ActualValue = $null } -Negative $false -Clause @{}
        $row = $script:HistoryItems[0]
        $row.Time | Should -Match '^\d{2}:\d{2}:\d{2}$'
        $row.Type | Should -Be 'File'
        $row.Expected | Should -Be '5'
        $row.Found | Should -Be ''
        $row.Target | Should -Be 'T'
    }

    It 'logs the test once' {
        & $script:DTPostTestResult -Result @{ Type = 'File'; Detected = $true; Details = ''; Target = 'T'; ExpectedValue = ''; ActualValue = '' } -Negative $true -Clause @{}
        Should -Invoke Write-Log -Times 1 -Exactly -ParameterFilter { $Message -like 'Test: type=File mode=Negative result=Fail*' }
    }
}

Describe 'DTRunTest' {
    BeforeEach {
        $global:Fx.Calls.Clear()
        Mock Write-Log {}
        Mock Test-RegistryKeyValueDetection { [void]$global:Fx.Calls.Add(@{ Name = 'Test-RegistryKeyValueDetection'; Args = $PesterBoundParameters }); @{ Type = 'Mocked'; Detected = $true; Details = 'mock details'; Target = 'mock'; ExpectedValue = 'e'; ActualValue = 'a' } }
        Mock Test-RegistryKeyDetection { [void]$global:Fx.Calls.Add(@{ Name = 'Test-RegistryKeyDetection'; Args = $PesterBoundParameters }); @{ Type = 'Mocked'; Detected = $true; Details = 'mock details'; Target = 'mock'; ExpectedValue = 'e'; ActualValue = 'a' } }
        Mock Test-FileDetection { [void]$global:Fx.Calls.Add(@{ Name = 'Test-FileDetection'; Args = $PesterBoundParameters }); @{ Type = 'Mocked'; Detected = $true; Details = 'mock details'; Target = 'mock'; ExpectedValue = 'e'; ActualValue = 'a' } }
        Mock Test-CompoundDetection { [void]$global:Fx.Calls.Add(@{ Name = 'Test-CompoundDetection'; Args = $PesterBoundParameters }); @{ Type = 'Mocked'; Detected = $true; Details = 'mock details'; Target = 'mock'; ExpectedValue = 'e'; ActualValue = 'a' } }
        $script:DTContext = New-Context
        $script:HistoryItems = New-Object 'System.Collections.ObjectModel.ObservableCollection[psobject]'
    }

    It 'stops with a message when the clause is not valid' {
        $script:DTContext.cboType.SelectedIndex = 1
        $script:DTContext.rkKeyPath.Text = ''
        & $script:DTRunTest
        $script:DTContext.txtResultBadge.Text | Should -Be 'Invalid clause'
        $script:DTContext.txtResultDetails.Text | Should -Be 'Key path is required.'
        $script:HistoryItems.Count | Should -Be 0
        $global:Fx.Calls.Count | Should -Be 0
    }

    It 'stops with a message when a compound clause has no data' {
        $script:DTContext.cboType.SelectedIndex = 4
        & $script:DTRunTest
        $script:DTContext.txtResultBadge.Text | Should -Be 'Invalid clause'
        $script:DTContext.txtResultDetails.Text | Should -Be 'No clause to evaluate.'
    }

    It 'runs a registry value test with the arguments of the clause' {
        $c = $script:DTContext
        $c.cboType.SelectedIndex = 0
        $c.rkvHiveHKCU.IsChecked = $true; $c.rkvView64.IsChecked = $false
        $c.rkvKeyPath.Text = 'Software\V'; $c.rkvValueName.Text = 'Version'; $c.rkvOperator.SelectedIndex = 1; $c.rkvExpected.Text = '3.0'
        & $script:DTRunTest
        $global:Fx.Calls.Count | Should -Be 1
        $a = $global:Fx.Calls[0]
        $a.Name | Should -Be 'Test-RegistryKeyValueDetection'
        $a.Args.RegistryKeyRelative | Should -Be 'Software\V'
        $a.Args.ValueName | Should -Be 'Version'
        $a.Args.ExpectedValue | Should -Be '3.0'
        $a.Args.Operator | Should -Be 'GreaterEquals'
        $a.Args.PropertyType | Should -Be 'String'
        $a.Args.Is64Bit | Should -BeFalse
        $a.Args.Hive | Should -Be 'CurrentUser'
        $script:HistoryItems.Count | Should -Be 1
        $c.txtResultBadge.Text | Should -Be 'Pass'
    }

    It 'runs a registry key test with the arguments of the clause' {
        $c = $script:DTContext
        $c.cboType.SelectedIndex = 1
        $c.rkKeyPath.Text = 'SOFTWARE\V'; $c.rkHiveHKLM.IsChecked = $true; $c.rkView64.IsChecked = $true
        & $script:DTRunTest
        $a = $global:Fx.Calls[0]
        $a.Name | Should -Be 'Test-RegistryKeyDetection'
        $a.Args.RegistryKeyRelative | Should -Be 'SOFTWARE\V'
        $a.Args.Is64Bit | Should -BeTrue
        $a.Args.Hive | Should -Be 'LocalMachine'
    }

    It 'runs a file test with the arguments of the clause' {
        $c = $script:DTContext
        $c.cboType.SelectedIndex = 2
        $c.fileFilePath.Text = 'C:\P'; $c.fileFileName.Text = 'p.exe'; $c.filePropType.SelectedIndex = 1; $c.fileOperator.SelectedIndex = 0; $c.fileExpected.Text = '1.2'; $c.fileView64.IsChecked = $false
        & $script:DTRunTest
        $a = $global:Fx.Calls[0]
        $a.Name | Should -Be 'Test-FileDetection'
        $a.Args.FilePath | Should -Be 'C:\P'
        $a.Args.FileName | Should -Be 'p.exe'
        $a.Args.PropertyType | Should -Be 'Version'
        $a.Args.Operator | Should -Be 'GreaterEquals'
        $a.Args.ExpectedValue | Should -Be '1.2'
        $a.Args.Is64Bit | Should -BeFalse
    }

    It 'runs a compound test with the connector and the clauses' {
        $c = $script:DTContext
        $c.cboType.SelectedIndex = 4
        $c.compoundConnector.Text = 'And'
        $c.CompoundClauseData = @(@{ Type = 'RegistryKey' }, @{ Type = 'File' })
        & $script:DTRunTest
        $a = $global:Fx.Calls[0]
        $a.Name | Should -Be 'Test-CompoundDetection'
        $a.Args.Connector | Should -Be 'And'
        $a.Args.Clauses.Count | Should -Be 2
    }

    It 'turns the result around in negative mode' {
        $script:DTContext.cboType.SelectedIndex = 1
        $script:DTContext.rkKeyPath.Text = 'SOFTWARE\V'
        $script:DTContext.chkNegativeMode.IsChecked = $true
        & $script:DTRunTest
        $script:DTContext.txtResultBadge.Text | Should -Be 'Fail'
        $script:HistoryItems[0].Mode | Should -Be 'Negative'
    }
}

Describe 'DTImportManifest' {
    BeforeEach {
        Mock Write-Log {}
        $script:DTContext = New-Context
        $script:DTContext.txtResultBadge.Text = 'old badge'; $script:DTContext.txtResultDetails.Text = 'old details'
        $script:window = $null
        $global:Fx.Dialog = [pscustomobject]@{ Filter = ''; Title = ''; FileName = 'C:\m\stage-manifest.json'; Answer = $true }
        Add-Member -InputObject $global:Fx.Dialog -MemberType ScriptMethod -Name ShowDialog -Value { param($owner) $this.Answer }
        # Newer Pester versions throw when a call matches no filter, so every other New-Object call goes to the real command.
        Mock New-Object { & (Get-Command -Name New-Object -CommandType Cmdlet) @PesterBoundParameters }
        Mock New-Object { $global:Fx.Dialog } -ParameterFilter { $TypeName -eq 'Microsoft.Win32.OpenFileDialog' }
        $global:Fx.Manifest = $null
        Mock Import-DetectionManifest { if ($global:Fx.Manifest -is [string]) { throw $global:Fx.Manifest }; $global:Fx.Manifest }
    }

    It 'changes nothing when the user cancels the dialog' {
        $global:Fx.Dialog.Answer = $false
        & $script:DTImportManifest
        Should -Invoke Import-DetectionManifest -Times 0 -Exactly
        $script:DTContext.txtResultBadge.Text | Should -Be 'old badge'
    }

    It 'asks for a JSON file' {
        $global:Fx.Dialog.Answer = $false
        & $script:DTImportManifest
        $global:Fx.Dialog.Filter | Should -BeLike '*(*.json)|*.json*'
        $global:Fx.Dialog.Title | Should -Be 'Import stage-manifest.json'
    }

    It 'shows the file name and the reason when the import fails' {
        $global:Fx.Manifest = "No 'Detection' block found in manifest"
        & $script:DTImportManifest
        $script:DTContext.txtResultBadge.Text | Should -Be 'Import failed'
        $script:DTContext.txtResultDetails.Text | Should -Be "Failed to import 'C:\m\stage-manifest.json': No 'Detection' block found in manifest"
        $script:DTContext.cboType.SelectedIndex | Should -Be 0
    }

    It 'fills the registry value panel and clears the old result' {
        $global:Fx.Manifest = @{ Type = 'RegistryKeyValue'; RegistryKeyRelative = 'Software\V'; ValueName = 'Ver'; ExpectedValue = '1.5'; Operator = 'LessThan'; Hive = 'CurrentUser'; Is64Bit = $false }
        & $script:DTImportManifest
        $c = $script:DTContext
        $c.cboType.SelectedIndex | Should -Be 0
        $c.rkvKeyPath.Text | Should -Be 'Software\V'
        $c.rkvValueName.Text | Should -Be 'Ver'
        $c.rkvExpected.Text | Should -Be '1.5'
        $c.rkvOperator.SelectedIndex | Should -Be 4
        $c.rkvHiveHKCU.IsChecked | Should -BeTrue
        $c.rkvHiveHKLM.IsChecked | Should -BeFalse
        $c.rkvView64.IsChecked | Should -BeFalse
        $c.rkvView32.IsChecked | Should -BeTrue
        $c.txtResultBadge.Text | Should -Be ''
        $c.txtResultDetails.Text | Should -Be ''
    }

    It 'fills the registry key panel' {
        $global:Fx.Manifest = @{ Type = 'RegistryKey'; RegistryKeyRelative = 'SOFTWARE\K'; Hive = 'LocalMachine'; Is64Bit = $true }
        & $script:DTImportManifest
        $c = $script:DTContext
        $c.cboType.SelectedIndex | Should -Be 1
        $c.rkKeyPath.Text | Should -Be 'SOFTWARE\K'
        $c.rkHiveHKLM.IsChecked | Should -BeTrue
        $c.rkView64.IsChecked | Should -BeTrue
    }

    It 'fills the file panel' {
        $global:Fx.Manifest = @{ Type = 'File'; FilePath = 'C:\P'; FileName = 'p.exe'; PropertyType = 'Version'; Operator = 'GreaterThan'; ExpectedValue = '2.0'; Is64Bit = $true }
        & $script:DTImportManifest
        $c = $script:DTContext
        $c.cboType.SelectedIndex | Should -Be 2
        $c.fileFilePath.Text | Should -Be 'C:\P'
        $c.fileFileName.Text | Should -Be 'p.exe'
        $c.filePropType.SelectedIndex | Should -Be 1
        $c.fileOperator.SelectedIndex | Should -Be 1
        $c.fileExpected.Text | Should -Be '2.0'
    }

    It 'fills the script panel' {
        $global:Fx.Manifest = @{ Type = 'Script'; ScriptText = '"found"'; ScriptLanguage = 'PowerShell' }
        & $script:DTImportManifest
        $script:DTContext.cboType.SelectedIndex | Should -Be 3
        $script:DTContext.scriptText.Text | Should -Be '"found"'
    }

    It 'lists the clauses of a compound manifest and keeps their data' {
        $long = 'x' * 100
        $global:Fx.Manifest = @{ Type = 'Compound'; Connector = 'Or'; Clauses = @(
                @{ Type = 'RegistryKeyValue'; Hive = 'LocalMachine'; RegistryKeyRelative = 'SOFTWARE\V'; ValueName = 'v'; Operator = 'IsEquals'; ExpectedValue = '1' }
                @{ Type = 'RegistryKey'; Hive = 'CurrentUser'; RegistryKeyRelative = 'Software\K' }
                @{ Type = 'File'; FilePath = 'C:\P'; FileName = 'p.exe'; PropertyType = 'Existence'; Operator = 'GreaterEquals'; ExpectedValue = '' }
                @{ Type = 'Script'; ScriptText = $long }
                @{ Type = 'Script'; ScriptText = 'short' }
                @{ Type = 'WmiQuery' }
            ) }
        & $script:DTImportManifest
        $c = $script:DTContext
        $c.cboType.SelectedIndex | Should -Be 4
        $c.compoundConnector.Text | Should -Be 'Or'
        $c.compoundClauses.Items.Count | Should -Be 6
        $c.compoundClauses.Items[0] | Should -Be "RKV:  LocalMachine\SOFTWARE\V [v] IsEquals '1'"
        $c.compoundClauses.Items[1] | Should -Be 'RK:   CurrentUser\Software\K'
        $c.compoundClauses.Items[2] | Should -Be "FILE: C:\P\p.exe (Existence GreaterEquals '')"
        $c.compoundClauses.Items[3] | Should -Be ('SCRIPT: ' + ('x' * 60) + '...')
        $c.compoundClauses.Items[4] | Should -Be 'SCRIPT: short...'
        $c.compoundClauses.Items[5] | Should -Be '?: WmiQuery'
        $c.CompoundClauseData.Count | Should -Be 6
        $c.txtResultBadge.Text | Should -Be ''
    }

    It 'replaces the list when a second compound manifest is imported' {
        $global:Fx.Manifest = @{ Type = 'Compound'; Connector = 'And'; Clauses = @(@{ Type = 'RegistryKey'; Hive = 'LocalMachine'; RegistryKeyRelative = 'A' }, @{ Type = 'RegistryKey'; Hive = 'LocalMachine'; RegistryKeyRelative = 'B' }) }
        & $script:DTImportManifest
        $global:Fx.Manifest = @{ Type = 'Compound'; Connector = 'And'; Clauses = @(@{ Type = 'RegistryKey'; Hive = 'LocalMachine'; RegistryKeyRelative = 'C' }) }
        & $script:DTImportManifest
        $script:DTContext.compoundClauses.Items.Count | Should -Be 1
        @($script:DTContext.CompoundClauseData).Count | Should -Be 1
    }

    It 'logs the import with the type and the file' {
        $global:Fx.Manifest = @{ Type = 'RegistryKey'; RegistryKeyRelative = 'K'; Hive = 'LocalMachine'; Is64Bit = $true }
        & $script:DTImportManifest
        Should -Invoke Write-Log -Times 1 -Exactly -ParameterFilter { $Message -eq 'Imported manifest: type=RegistryKey from C:\m\stage-manifest.json' }
    }
}

Describe 'DTRepopulateFromHistory' {
    BeforeEach {
        Mock Write-Log {}
        $script:DTContext = New-Context
    }

    It 'changes nothing when no row is selected' {
        $script:DTContext.cboType.SelectedIndex = 2
        & $script:DTRepopulateFromHistory
        $script:DTContext.cboType.SelectedIndex | Should -Be 2
    }

    It 'only switches the type for a row without a clause snapshot: <Type>' -ForEach @(
        @{ Type = 'RegistryKeyValue'; Index = 0 }, @{ Type = 'RegistryKey'; Index = 1 }, @{ Type = 'File'; Index = 2 }, @{ Type = 'Script'; Index = 3 }, @{ Type = 'Compound'; Index = 4 }
    ) {
        $script:DTContext.dgHistory.SelectedItem = [pscustomobject]@{ Type = $Type; Clause = $null; Negative = $false }
        $script:DTContext.cboType.SelectedIndex = -1
        & $script:DTRepopulateFromHistory
        $script:DTContext.cboType.SelectedIndex | Should -Be $Index
    }

    It 'restores a registry value clause and the negative mode' {
        $clause = @{ Type = 'RegistryKeyValue'; Hive = 'CurrentUser'; Is64Bit = $false; RegistryKeyRelative = 'Software\V'; ValueName = 'Ver'; Operator = 'GreaterThan'; ExpectedValue = '1.0' }
        $script:DTContext.dgHistory.SelectedItem = [pscustomobject]@{ Type = 'RegistryKeyValue'; Clause = $clause; Negative = $true }
        & $script:DTRepopulateFromHistory
        $c = $script:DTContext
        $c.chkNegativeMode.IsChecked | Should -BeTrue
        $c.cboType.SelectedIndex | Should -Be 0
        $c.rkvKeyPath.Text | Should -Be 'Software\V'
        $c.rkvValueName.Text | Should -Be 'Ver'
        $c.rkvExpected.Text | Should -Be '1.0'
        $c.rkvOperator.SelectedIndex | Should -Be 2
        $c.rkvHiveHKCU.IsChecked | Should -BeTrue
        $c.rkvView32.IsChecked | Should -BeTrue
    }

    It 'restores a registry key clause' {
        $script:DTContext.dgHistory.SelectedItem = [pscustomobject]@{ Type = 'RegistryKey'; Negative = $false; Clause = @{ Type = 'RegistryKey'; Hive = 'LocalMachine'; Is64Bit = $true; RegistryKeyRelative = 'SOFTWARE\K' } }
        & $script:DTRepopulateFromHistory
        $script:DTContext.cboType.SelectedIndex | Should -Be 1
        $script:DTContext.rkKeyPath.Text | Should -Be 'SOFTWARE\K'
        $script:DTContext.rkHiveHKLM.IsChecked | Should -BeTrue
        $script:DTContext.rkView64.IsChecked | Should -BeTrue
    }

    It 'restores a file clause' {
        $script:DTContext.dgHistory.SelectedItem = [pscustomobject]@{ Type = 'File'; Negative = $false; Clause = @{ Type = 'File'; FilePath = 'C:\P'; FileName = 'p.exe'; PropertyType = 'Version'; Operator = 'LessEquals'; ExpectedValue = '9'; Is64Bit = $false } }
        & $script:DTRepopulateFromHistory
        $c = $script:DTContext
        $c.cboType.SelectedIndex | Should -Be 2
        $c.fileFilePath.Text | Should -Be 'C:\P'
        $c.filePropType.SelectedIndex | Should -Be 1
        $c.fileOperator.SelectedIndex | Should -Be 3
        $c.fileView32.IsChecked | Should -BeTrue
    }

    It 'restores a script clause' {
        $script:DTContext.dgHistory.SelectedItem = [pscustomobject]@{ Type = 'Script'; Negative = $false; Clause = @{ Type = 'Script'; ScriptText = '"x"' } }
        & $script:DTRepopulateFromHistory
        $script:DTContext.cboType.SelectedIndex | Should -Be 3
        $script:DTContext.scriptText.Text | Should -Be '"x"'
    }

    It 'restores a compound clause with its list and its data' {
        $clauses = @(@{ Type = 'RegistryKey'; RegistryKeyRelative = 'A' }, @{ Type = 'File'; FileName = 'b' })
        $script:DTContext.dgHistory.SelectedItem = [pscustomobject]@{ Type = 'Compound'; Negative = $false; Clause = @{ Type = 'Compound'; Connector = 'And'; Clauses = $clauses } }
        & $script:DTRepopulateFromHistory
        $c = $script:DTContext
        $c.cboType.SelectedIndex | Should -Be 4
        $c.compoundConnector.Text | Should -Be 'And'
        $c.compoundClauses.Items.Count | Should -Be 2
        $c.compoundClauses.Items[0] | Should -BeLike 'RegistryKey: {*'
        $c.CompoundClauseData.Count | Should -Be 2
    }
}

Describe 'History export and clear' {
    BeforeEach {
        Mock Write-Log {}
        $script:DTContext = New-Context
        $script:window = $null
        $script:reportsDir = $TestDrive
        $script:HistoryItems = New-Object 'System.Collections.ObjectModel.ObservableCollection[psobject]'
        $global:Fx.Calls.Clear()
        $global:Fx.Dialog = [pscustomobject]@{ Filter = ''; InitialDirectory = ''; FileName = ''; Answer = $true; Chosen = (Join-Path $TestDrive 'history.out') }
        Add-Member -InputObject $global:Fx.Dialog -MemberType ScriptMethod -Name ShowDialog -Value { param($owner) if ($this.Answer) { $this.FileName = $this.Chosen }; $this.Answer }
        # Newer Pester versions throw when a call matches no filter, so every other New-Object call goes to the real command.
        Mock New-Object { & (Get-Command -Name New-Object -CommandType Cmdlet) @PesterBoundParameters }
        Mock New-Object { $global:Fx.Dialog } -ParameterFilter { $TypeName -eq 'Microsoft.Win32.SaveFileDialog' }
        Mock Export-DetectionResultsCsv { [void]$global:Fx.Calls.Add(@{ Name = 'Csv'; Rows = $DataTable.Rows.Count; Path = $OutputPath; IsTable = ($DataTable -is [System.Data.DataTable]) }) }
        Mock Export-DetectionResultsHtml { [void]$global:Fx.Calls.Add(@{ Name = 'Html'; Rows = $DataTable.Rows.Count; Path = $OutputPath; IsTable = ($DataTable -is [System.Data.DataTable]) }) }
        $script:HistoryItems.Add([pscustomobject]@{ Time = 't'; Type = 'File'; Mode = 'Positive'; Target = 'a'; Expected = 'b'; Found = 'c'; Result = 'Pass' })
    }

    It 'exports the history as <Kind> to the chosen file' -ForEach @(
        @{ Kind = 'CSV'; Block = 'DTExportHistoryCsv'; Call = 'Csv'; Filter = 'CSV (*.csv)|*.csv'; Prefix = 'detection-history-' }
        @{ Kind = 'HTML'; Block = 'DTExportHistoryHtml'; Call = 'Html'; Filter = 'HTML (*.html)|*.html'; Prefix = 'detection-history-' }
    ) {
        & $script:Blocks[$Block]
        $global:Fx.Calls.Count | Should -Be 1
        $global:Fx.Calls[0].Name | Should -Be $Call
        $global:Fx.Calls[0].Rows | Should -Be 1
        $global:Fx.Calls[0].IsTable | Should -BeTrue
        $global:Fx.Calls[0].Path | Should -Be $global:Fx.Dialog.Chosen
        $global:Fx.Dialog.Filter | Should -Be $Filter
        $global:Fx.Dialog.InitialDirectory | Should -Be $TestDrive
    }

    It 'names the default file with the date, for <Block>' -ForEach @(@{ Block = 'DTExportHistoryCsv'; Ext = 'csv' }, @{ Block = 'DTExportHistoryHtml'; Ext = 'html' }) {
        $global:Fx.Dialog.Answer = $false
        # The dialog gets its default name before it is shown; read it from the object after the call.
        & $script:Blocks[$Block]
        $global:Fx.Dialog.FileName | Should -Match ('^detection-history-\d{8}-\d{6}\.' + $Ext + '$')
    }

    It 'does not open a dialog or export when the history is empty, for <Block>' -ForEach @(@{ Block = 'DTExportHistoryCsv' }, @{ Block = 'DTExportHistoryHtml' }) {
        $script:HistoryItems.Clear()
        & $script:Blocks[$Block]
        $global:Fx.Calls.Count | Should -Be 0
        Should -Invoke New-Object -Times 0 -Exactly -ParameterFilter { $TypeName -eq 'Microsoft.Win32.SaveFileDialog' }
    }

    It 'does not export when the user cancels, for <Block>' -ForEach @(@{ Block = 'DTExportHistoryCsv' }, @{ Block = 'DTExportHistoryHtml' }) {
        $global:Fx.Dialog.Answer = $false
        & $script:Blocks[$Block]
        $global:Fx.Calls.Count | Should -Be 0
    }

    It 'shows the reason in the result panel when the export fails, for <Block>' -ForEach @(@{ Block = 'DTExportHistoryCsv'; Fn = 'Export-DetectionResultsCsv' }, @{ Block = 'DTExportHistoryHtml'; Fn = 'Export-DetectionResultsHtml' }) {
        Mock $Fn { throw 'disk is full' }
        { & $script:Blocks[$Block] } | Should -Not -Throw
        $script:DTContext.txtResultBadge.Text | Should -Be 'Export failed'
        $script:DTContext.txtResultDetails.Text | Should -BeLike '*disk is full*'
        $script:DTContext.txtResultDetails.Text | Should -BeLike ('*' + $global:Fx.Dialog.Chosen + '*')
    }

    It 'empties the history' {
        & $script:DTClearHistory
        $script:HistoryItems.Count | Should -Be 0
    }
}

Describe 'Installed applications list' {
    BeforeEach {
        Mock Write-Log {}
        $script:IAItems = New-Object 'System.Collections.ObjectModel.ObservableCollection[psobject]'
        $apps = @(
            [pscustomobject]@{ DisplayName = 'Alpha Tool'; Publisher = 'Acme'; DisplayVersion = '1.0'; Scope = 'Machine'; Architecture = 'x64' }
            [pscustomobject]@{ DisplayName = 'Beta (x86)'; Publisher = 'Contoso'; DisplayVersion = '2.5'; Scope = 'Machine'; Architecture = 'x86' }
            [pscustomobject]@{ DisplayName = 'Gamma+'; Publisher = 'Acme'; DisplayVersion = '3.0'; Scope = 'User'; Architecture = 'x64' }
            [pscustomobject]@{ DisplayName = 'Delta'; Publisher = ''; DisplayVersion = ''; Scope = 'User'; Architecture = 'x64' }
        )
        $script:IAContext = @{ AllApps = $apps; txtFilter = [pscustomobject]@{ Text = '' }; cboScope = New-Combo -Item 'All', 'Machine (HKLM)', 'User (HKCU)' }
    }

    It 'shows <Expected> for filter "<Filter>" and scope index <Scope>' -ForEach @(
        @{ Filter = '';        Scope = 0; Expected = 'Alpha Tool,Beta (x86),Gamma+,Delta' }
        @{ Filter = 'alpha';   Scope = 0; Expected = 'Alpha Tool' }
        @{ Filter = 'ACME';    Scope = 0; Expected = 'Alpha Tool,Gamma+' }
        @{ Filter = '2.5';     Scope = 0; Expected = 'Beta (x86)' }
        @{ Filter = '(x86)';   Scope = 0; Expected = 'Beta (x86)' }
        @{ Filter = '+';       Scope = 0; Expected = 'Gamma+' }
        @{ Filter = '.*';      Scope = 0; Expected = '' }
        @{ Filter = '[';       Scope = 0; Expected = '' }
        @{ Filter = 'nothing'; Scope = 0; Expected = '' }
        @{ Filter = '';        Scope = 1; Expected = 'Alpha Tool,Beta (x86)' }
        @{ Filter = '';        Scope = 2; Expected = 'Gamma+,Delta' }
        @{ Filter = 'acme';    Scope = 2; Expected = 'Gamma+' }
        @{ Filter = 'acme';    Scope = 1; Expected = 'Alpha Tool' }
    ) {
        $script:IAContext.txtFilter.Text = $Filter
        $script:IAContext.cboScope.SelectedIndex = $Scope
        & $script:IAApplyFilter
        ($script:IAItems.DisplayName -join ',') | Should -Be $Expected
    }

    It 'leaves the list alone before the applications are loaded' {
        $script:IAContext.AllApps = $null
        $script:IAItems.Add([pscustomobject]@{ DisplayName = 'keep' })
        & $script:IAApplyFilter
        $script:IAItems.Count | Should -Be 1
    }

    It 'loads the applications and then filters them' {
        Mock Get-InstalledApplications { @([pscustomobject]@{ DisplayName = 'Loaded'; Publisher = ''; DisplayVersion = ''; Scope = 'Machine'; Architecture = 'x64' }) }
        & $script:IARefresh
        @($script:IAContext.AllApps).Count | Should -Be 1
        $script:IAItems.Count | Should -Be 1
        Should -Invoke Get-InstalledApplications -Times 1 -Exactly
    }
}

Describe 'IAOnAppSelected' {
    BeforeEach {
        $fields = 'dtlDisplayName', 'dtlPublisher', 'dtlDisplayVersion', 'dtlArchScope', 'dtlRegistryKey', 'dtlUninstallString', 'dtlQuietUninstall', 'dtlInstallLocation', 'dtlInstallDate'
        $script:IAContext = @{ dgApps = [pscustomobject]@{ SelectedItem = $null }; btnUseForDetection = [pscustomobject]@{ Visibility = 'Collapsed' }; btnCopyDetails = [pscustomobject]@{ Visibility = 'Collapsed' } }
        foreach ($f in $fields) { $script:IAContext[$f] = [pscustomobject]@{ Text = 'stale' } }
    }

    It 'shows the buttons and the details of the selected application' {
        $script:IAContext.dgApps.SelectedItem = [pscustomobject]@{ DisplayName = 'App'; Publisher = 'Pub'; DisplayVersion = '1.0'; Architecture = 'x86'; Scope = 'Machine'; RegistryKey = 'HKLM\K'; UninstallString = 'u'; QuietUninstallString = 'q'; InstallLocation = 'C:\L'; InstallDate = '20260101' }
        & $script:IAOnAppSelected
        $c = $script:IAContext
        $c.btnUseForDetection.Visibility | Should -Be 'Visible'
        $c.btnCopyDetails.Visibility | Should -Be 'Visible'
        $c.dtlDisplayName.Text | Should -Be 'App'
        $c.dtlArchScope.Text | Should -Be 'x86 / Machine'
        $c.dtlRegistryKey.Text | Should -Be 'HKLM\K'
        $c.dtlQuietUninstall.Text | Should -Be 'q'
        $c.dtlInstallDate.Text | Should -Be '20260101'
    }

    It 'hides the buttons and clears the details when nothing is selected' {
        $script:IAContext.btnUseForDetection.Visibility = 'Visible'; $script:IAContext.btnCopyDetails.Visibility = 'Visible'
        & $script:IAOnAppSelected
        $script:IAContext.btnUseForDetection.Visibility | Should -Be 'Collapsed'
        $script:IAContext.btnCopyDetails.Visibility | Should -Be 'Collapsed'
        foreach ($f in 'dtlDisplayName', 'dtlRegistryKey', 'dtlInstallDate') { $script:IAContext[$f].Text | Should -Be '' }
    }
}

Describe 'IAUseForDetection' {
    BeforeEach {
        Mock Write-Log {}
        $script:DTContext = New-Context
        $script:DTContext.txtResultBadge.Text = 'old'; $script:DTContext.txtResultDetails.Text = 'old'
        $script:DTContext.rkvOperator.SelectedIndex = 3
        $script:IAContext = @{ dgApps = [pscustomobject]@{ SelectedItem = $null } }
        $global:Fx.Shown = New-Object System.Collections.ArrayList
        $script:ShowModule = { param($Name) [void]$global:Fx.Shown.Add($Name) }
    }
    AfterEach {
        # Put the real block back for the tests of the other Describe blocks.
        $script:ShowModule = [scriptblock]::Create((Get-ShellBlockText -Name 'ShowModule'))
    }

    It 'does nothing when no application is selected' {
        & $script:IAUseForDetection
        $global:Fx.Shown.Count | Should -Be 0
    }

    It 'opens the Detection Tester and fills the registry value panel for <Name>' -ForEach @(
        @{ Name = 'a 64-bit machine application'; Hive = 'LocalMachine'; Arch = 'x64'; HKCU = $false; V64 = $true }
        @{ Name = 'a 32-bit machine application'; Hive = 'LocalMachine'; Arch = 'x86'; HKCU = $false; V64 = $false }
        @{ Name = 'a user application'; Hive = 'CurrentUser'; Arch = 'x64'; HKCU = $true; V64 = $true }
    ) {
        $script:IAContext.dgApps.SelectedItem = [pscustomobject]@{ DisplayName = 'App'; RegistryKey = 'HKLM\K'; RegistryKeyRelative = 'SOFTWARE\K'; DisplayVersion = '4.5'; Hive = $Hive; Architecture = $Arch }
        $script:DTContext.cboType.SelectedIndex = 3
        & $script:IAUseForDetection
        $global:Fx.Shown | Should -Be @('DetectionTester')
        $c = $script:DTContext
        $c.cboType.SelectedIndex | Should -Be 0
        $c.rkvKeyPath.Text | Should -Be 'SOFTWARE\K'
        $c.rkvValueName.Text | Should -Be 'DisplayVersion'
        $c.rkvExpected.Text | Should -Be '4.5'
        $c.rkvHiveHKCU.IsChecked | Should -Be $HKCU
        $c.rkvHiveHKLM.IsChecked | Should -Be (-not $HKCU)
        $c.rkvView64.IsChecked | Should -Be $V64
        $c.rkvView32.IsChecked | Should -Be (-not $V64)
        $c.rkvOperator.SelectedIndex | Should -Be 0
        $c.txtResultBadge.Text | Should -Be ''
        $c.txtResultDetails.Text | Should -Be ''
    }
}

Describe 'Format-DTAppDetails' {
    BeforeAll { . ([scriptblock]::Create((Get-ShellFunctionText -Name 'Format-DTAppDetails'))) }

    It 'lists every field of the application with its name' {
        $a = [pscustomobject]@{ DisplayName = 'App'; Publisher = 'Pub'; DisplayVersion = '1.0'; Architecture = 'x64'; Scope = 'Machine'; RegistryKey = 'HKLM\K'; UninstallString = 'u'; QuietUninstallString = 'q'; InstallLocation = 'C:\L'; InstallDate = '20260101' }
        $text = Format-DTAppDetails -App $a
        $lines = $text -split "`r`n"
        $lines.Count | Should -Be 10
        $lines[0] | Should -Be 'DisplayName:          App'
        $lines[5] | Should -Be 'RegistryKey:          HKLM\K'
        $lines[9] | Should -Be 'InstallDate:          20260101'
    }

    It 'gives an empty value for a field that the application does not have' {
        (Format-DTAppDetails -App ([pscustomobject]@{ DisplayName = 'App' })) | Should -BeLike "DisplayName:          App`r`nPublisher:            `r`n*"
    }
}

Describe 'ShowModule' {
    BeforeEach {
        $script:Prefs = @{ DarkMode = $true; ActiveModule = 'DetectionTester' }
        $script:DTRoot = $null; $script:IARoot = $null
        $global:Fx.Loads = New-Object System.Collections.ArrayList
        $script:LoadDetectionTester = { [void]$global:Fx.Loads.Add('DT'); 'dt-root' }
        $script:LoadInstalledApps = { [void]$global:Fx.Loads.Add('IA'); 'ia-root' }
        $contentHost = [pscustomobject]@{ Content = $null }
        $txtModuleTitle = [pscustomobject]@{ Text = '' }
        $txtModuleSubtitle = [pscustomobject]@{ Text = '' }
        $txtStatus = [pscustomobject]@{ Text = '' }
        $script:Show = $script:ShowModule
    }

    It 'shows the <Name> module with its title and status' -ForEach @(
        @{ Name = 'DetectionTester'; Root = 'dt-root'; Title = 'Detection Tester'; Status = 'Detection Tester selected.' }
        @{ Name = 'InstalledApps'; Root = 'ia-root'; Title = 'Installed Applications'; Status = 'Installed Applications selected.' }
    ) {
        & $script:ShowModule $Name
        $contentHost.Content | Should -Be $Root
        $txtModuleTitle.Text | Should -Be $Title
        $txtModuleSubtitle.Text | Should -Not -BeNullOrEmpty
        $txtStatus.Text | Should -Be $Status
        $script:Prefs.ActiveModule | Should -Be $Name
    }

    It 'loads a module once and shows the same root the second time' {
        & $script:ShowModule 'InstalledApps'
        & $script:ShowModule 'DetectionTester'
        & $script:ShowModule 'InstalledApps'
        @($global:Fx.Loads | Where-Object { $_ -eq 'IA' }).Count | Should -Be 1
        $contentHost.Content | Should -Be 'ia-root'
    }

    It 'shows the Detection Tester for a name it does not know' {
        & $script:ShowModule 'Unknown'
        $contentHost.Content | Should -Be 'dt-root'
        $script:Prefs.ActiveModule | Should -Be 'DetectionTester'
    }
}

Describe 'OnClosing' {
    It 'saves the preferences and the window state and logs' {
        Mock Write-Log {}
        $script:Prefs = @{ DarkMode = $false }
        $script:wsPath = Join-Path $TestDrive 'closing-ws.json'
        $script:prefsPath = Join-Path $TestDrive 'closing-prefs.json'
        $script:window = [pscustomobject]@{ WindowState = [System.Windows.WindowState]::Normal; Left = 1; Top = 2; ActualWidth = 3; ActualHeight = 4; RestoreBounds = $null }
        & $script:OnClosing
        Test-Path -LiteralPath $script:wsPath | Should -BeTrue
        (Get-Content -LiteralPath $script:prefsPath -Raw | ConvertFrom-Json).DarkMode | Should -BeFalse
        Should -Invoke Write-Log -Times 1 -Exactly -ParameterFilter { $Message -eq 'Detection Tester WPF shell closed.' }
    }
}

Describe 'Start script source' {
    It 'never stores a value under a key that differs from a control name only in case' {
        $ast = Get-ShellAst
        $loader = $ast.Find({ param($n) $n -is [System.Management.Automation.Language.AssignmentStatementAst] -and $n.Left.Extent.Text -eq '$script:LoadDetectionTester' }, $true)
        $names = @($loader.Right.FindAll({ param($n) $n -is [System.Management.Automation.Language.ArrayLiteralAst] }, $true) | Select-Object -First 1 | ForEach-Object { $_.Elements.Value })
        $names.Count | Should -BeGreaterThan 20
        $assigned = @($ast.FindAll({
                    param($n)
                    $n -is [System.Management.Automation.Language.AssignmentStatementAst] -and
                    $n.Left -is [System.Management.Automation.Language.MemberExpressionAst] -and
                    $n.Left.Expression.Extent.Text -in '$ctx', '$dt', '$script:DTContext', '$script:IAContext'
                }, $true) | ForEach-Object { $_.Left.Member.Value } | Sort-Object -Unique)
        $assigned.Count | Should -BeGreaterThan 0
        $clash = @($assigned | Where-Object { $a = $_; $names | Where-Object { $_ -ieq $a } })
        $clash | Should -BeNullOrEmpty
    }

    It 'reads the compound clause control only through its Items and the data only through CompoundClauseData' {
        $ast = Get-ShellAst
        $bare = @($ast.FindAll({
                    param($n)
                    $n -is [System.Management.Automation.Language.MemberExpressionAst] -and $n.Member.Value -eq 'compoundClauses' -and
                    -not ($n.Parent -is [System.Management.Automation.Language.MemberExpressionAst] -and $n.Parent.Member.Value -eq 'Items')
                }, $true))
        $bare | Should -BeNullOrEmpty
        @($ast.FindAll({ param($n) $n -is [System.Management.Automation.Language.MemberExpressionAst] -and $n.Member.Value -eq 'CompoundClauseData' }, $true)).Count | Should -BeGreaterOrEqual 3
    }

    It 'clears the previous result when a clause is filled from an application or a manifest' -ForEach @(@{ Name = 'IAUseForDetection' }, @{ Name = 'DTImportManifest' }) {
        $text = Get-ShellBlockText -Name $Name
        $text | Should -Match "txtResultBadge\.Text\s*=\s*''"
        $text | Should -Match "txtResultDetails\.Text\s*=\s*''"
    }

    It 'gives every fixed-width grid column nine pixels for each character of its header' -ForEach @(@{ File = 'DetectionTester.xaml' }, @{ File = 'InstalledApps.xaml' }) {
        $xaml = Get-Content -LiteralPath (Join-Path $script:AppRoot "Modules\$File") -Raw
        $columns = @([regex]::Matches($xaml, '<DataGridTextColumn\s+Header="([^"]+)"[^>]*?Width="(\d+)"'))
        $columns.Count | Should -BeGreaterThan 3
        foreach ($m in $columns) {
            [int]$m.Groups[2].Value | Should -BeGreaterOrEqual (9 * $m.Groups[1].Value.Length) -Because "$File column $($m.Groups[1].Value)"
        }
    }

    It 'declares the same version in the start script, the module manifest and the changelog' {
        $shell = Get-Content -LiteralPath (Join-Path $script:AppRoot 'start-detectiontester.ps1') -TotalCount 80
        $shellVersion = ($shell | Where-Object { $_ -match '^\s*Version\s*:\s*([0-9][0-9\.]*[0-9])\s*$' } | Select-Object -First 1) -replace '^\s*Version\s*:\s*', ''
        $psd1 = Import-PowerShellDataFile -LiteralPath (Join-Path $script:AppRoot 'Module\DetectionTesterCommon.psd1')
        $changelog = Get-Content -LiteralPath (Join-Path $script:AppRoot 'CHANGELOG.md') | Where-Object { $_ -match '^## \[' } | Select-Object -First 1
        $shellVersion | Should -Not -BeNullOrEmpty
        $psd1.ModuleVersion | Should -Be $shellVersion
        $changelog.StartsWith("## [$shellVersion] - ") | Should -BeTrue -Because 'the top heading of the changelog names the version'
        $shellVersion | Should -Match '^\d{4}\.\d{2}\.\d{2}\.\d{4}$'
    }
}

Describe 'Get-ShellRuntimePart' {
    It 'throws when a listed function is missing from the start script' {
        { Get-ShellRuntimePart -FunctionName 'Get-DTComboValue', 'Zz-NoSuchFunction' } | Should -Throw '*function Zz-NoSuchFunction not found*'
    }
}