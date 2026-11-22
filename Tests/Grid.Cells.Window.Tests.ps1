#Requires -Modules @{ ModuleName = 'Pester'; ModuleVersion = '5.1.0' }
# Window test of the grid cells. It gives each real grid of the two modules, loaded into the content host of the main window,
# the rows of GridFixtures.ps1, through
# ItemsSource as the refresh handler does, and reads back what the grid shows: the row count and the text of each
# cell. A binding that does not resolve shows an empty cell and fails here.

BeforeDiscovery {
    . "$PSScriptRoot\WindowHarness.ps1"
    . "$PSScriptRoot\GridData.Helpers.ps1"
    $script:WindowGuard = Test-WindowTestsAllowed
    $script:GridCases = @(@{ Grid = 'skipped' })
    if ($script:WindowGuard.Allowed) {
        $script:GridCases = @(Get-XamlGridColumn -XamlPath (Get-ChildItem -LiteralPath (Join-Path (Split-Path -Parent $PSScriptRoot) 'Modules') -Filter '*.xaml' | ForEach-Object FullName) |
                Select-Object -ExpandProperty Grid -Unique | ForEach-Object { @{ Grid = $_ } })
    }
}

BeforeAll {
    . "$PSScriptRoot\TestHelpers.ps1"
    . "$PSScriptRoot\WindowHarness.ps1"
    $script:WindowGuard = Test-WindowTestsAllowed
    if (-not $script:WindowGuard.Allowed) { return }
    . "$PSScriptRoot\WindowConfig.ps1"
    . "$PSScriptRoot\GridData.Helpers.ps1"
    . "$PSScriptRoot\GridFixtures.ps1"
    Import-AppModule
    Mock Write-Log {} -ModuleName DetectionTesterCommon
    # The result handler of the script calls Write-Log from the script scope.
    Mock Write-Log {}
    $script:Fixture = New-GridFixture
    $script:Columns = @(Get-XamlGridColumn -XamlPath (Get-ChildItem -LiteralPath (Join-Path (Split-Path -Parent $PSScriptRoot) 'Modules') -Filter '*.xaml' | ForEach-Object FullName))
    $script:XamlOf = @{}
    foreach ($c in $script:Columns) { $script:XamlOf[$c.Grid] = $c.Xaml }
    $script:Shown = @{}
    foreach ($g in @($script:Fixture.Keys)) {
        $script:Shown[$g] = Invoke-WpfSta -Arguments @{ Xaml = $script:AppConfig.Xaml; Grid = $g; Rows = @($script:Fixture[$g]); ModuleXaml = $script:XamlOf[$g] } -Script {
            $w = $null
            try {
                $w = Open-TestWindow -XamlPath (Join-Path $Root $Arguments.Xaml)
                # The shell loads the module XAML into the content host.
                [xml]$x = Get-Content -LiteralPath $Arguments.ModuleXaml -Raw -Encoding UTF8
                $moduleRoot = [System.Windows.Markup.XamlReader]::Load((New-Object System.Xml.XmlNodeReader $x))
                $w.FindName('contentHost').Content = $moduleRoot
                Wait-Layout $w
                $grid = $moduleRoot.FindName($Arguments.Grid)
                $grid.ItemsSource = @($Arguments.Rows)
                Wait-Layout $w
                $cells = @(for ($i = 0; $i -lt $grid.Items.Count; $i++) {
                        foreach ($col in $grid.Columns) {
                            $content = $col.GetCellContent($grid.Items[$i])
                            [PSCustomObject]@{ Row = $i; Header = [string]$col.Header; Text = $(if ($content -is [System.Windows.Controls.TextBlock]) { $content.Text } else { $null }) }
                        }
                    })
                [PSCustomObject]@{ Count = $grid.Items.Count; Cells = $cells }
            }
            finally { if ($w) { $w.Close() } }
        }
    }
}

AfterAll { if (Get-Command Close-WpfSession -ErrorAction SilentlyContinue) { Close-WpfSession } }

Describe 'Grid cells' -Skip:(-not $script:WindowGuard.Allowed) {
    It 'shows every row of the grid <Grid>' -ForEach $script:GridCases {
        $shown = @($script:Shown[$Grid])[0]
        $shown.Count | Should -Be @($script:Fixture[$Grid]).Count
    }

    It 'shows the value of every bound cell of the grid <Grid>' -ForEach $script:GridCases {
        $shown = @($script:Shown[$Grid])[0]
        $rows = @($script:Fixture[$Grid])
        foreach ($column in @($script:Columns | Where-Object Grid -eq $Grid)) {
            for ($i = 0; $i -lt $rows.Count; $i++) {
                $rows[$i].PSObject.Properties.Name | Should -Contain $column.Path -Because "the grid $Grid binds a column to $($column.Path)"
                $cell = @($shown.Cells | Where-Object { $_.Row -eq $i -and $_.Header -eq $column.Header })[0]
                $cell.Text | Should -Be ([string]$rows[$i].($column.Path)) -Because "row $i of $($Grid), column $($column.Path)"
            }
        }
    }
}
