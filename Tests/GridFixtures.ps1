# Realistic rows for the grid data tests. The history rows come out of the real detection functions of the module and
# the real result handler of start-detectiontester.ps1 (DTPostTestResult, cut from its syntax tree as in
# Shell.Logic.Tests.ps1). The application rows come out of the real Get-InstalledApplications, which reads the
# registry of this machine, plus one throwaway entry under HKCU, and out of the real list filter (IAApplyFilter).
# No row is written by hand, so a property name that the code stops producing shows in the grid tests.
# Call it where Write-Log is mocked and the module is imported.

function New-GridFixture {
    Add-Type -AssemblyName PresentationFramework, PresentationCore, WindowsBase
    . ([scriptblock]::Create((Get-ShellFunctionText -Name 'Get-DTComboValue')))
    $postResult = [scriptblock]::Create((Get-ShellBlockText -Name 'DTPostTestResult'))
    $applyFilter = [scriptblock]::Create((Get-ShellBlockText -Name 'IAApplyFilter'))

    $work = Join-Path ([System.IO.Path]::GetTempPath()) ('dt-grid-' + [guid]::NewGuid().ToString('N'))
    New-Item -ItemType Directory -Path $work -Force | Out-Null
    $relative = $null
    try {
        [System.IO.File]::WriteAllText((Join-Path $work 'application.exe'), 'x')
        $relative = New-TestRegistryKey -Value @{ DisplayVersion = '1.2.3' }
        $arp = New-TestRegistryKey -Parent 'Software\Microsoft\Windows\CurrentVersion\Uninstall' -Value @{ DisplayName = 'Contoso Viewer'; Publisher = 'Contoso Ltd'; DisplayVersion = '4.2.0'; UninstallString = 'C:\Contoso\uninstall.exe' }

        # The history grid: one row for each result that the handler shows.
        $script:HistoryItems = New-Object 'System.Collections.ObjectModel.ObservableCollection[psobject]'
        $script:DTContext = @{ txtResultBadge = [pscustomobject]@{ Text = '' }; txtResultDetails = [pscustomobject]@{ Text = '' } }
        $results = @(
            @{ Negative = $false; Clause = @{ Type = 'RegistryKeyValue' }; Result = (Test-RegistryKeyValueDetection -RegistryKeyRelative $relative -ValueName 'DisplayVersion' -ExpectedValue '1.2.3' -Operator 'IsEquals' -PropertyType 'String' -Is64Bit $true -Hive 'CurrentUser') }
            @{ Negative = $true; Clause = @{ Type = 'RegistryKey' }; Result = (Test-RegistryKeyDetection -RegistryKeyRelative $relative -Is64Bit $true -Hive 'CurrentUser') }
            @{ Negative = $false; Clause = @{ Type = 'File' }; Result = (Test-FileDetection -FilePath $work -FileName 'application.exe' -PropertyType 'Existence' -Is64Bit $true) }
        )
        foreach ($r in $results) { $null = & $postResult -Result $r.Result -Negative $r.Negative -Clause $r.Clause }
        $history = @($script:HistoryItems)

        # The installed applications grid.
        $script:IAItems = New-Object 'System.Collections.ObjectModel.ObservableCollection[psobject]'
        $combo = [pscustomobject]@{ SelectedItem = 'All' }
        $script:IAContext = @{ AllApps = @(Get-InstalledApplications); txtFilter = [pscustomobject]@{ Text = '' }; cboScope = $combo }
        $null = & $applyFilter
        # A handful of rows: the grid draws only the rows that fit in the window, and the window test reads the cells of each row.
        $apps = @(@($script:IAItems | Where-Object DisplayName -eq 'Contoso Viewer') + @($script:IAItems | Where-Object DisplayName -ne 'Contoso Viewer' | Select-Object -First 5))

        @{ dgHistory = $history; dgApps = $apps }
    }
    finally {
        if ($relative) { Remove-TestRegistryKey -Relative $relative }
        if ($arp) { Remove-TestRegistryKey -Relative $arp }
        Remove-Item -LiteralPath $work -Recurse -Force -ErrorAction SilentlyContinue
    }
}

# Columns that a fixture cannot fill on any row, each with the reason. The grid test accepts only these.
function Get-GridEmptyAllowed {
    @{}
}
