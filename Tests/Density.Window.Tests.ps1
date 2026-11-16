#Requires -Modules @{ ModuleName = 'Pester'; ModuleVersion = '5.1.0' }
# Window test for the compact 1080p density and for the vacant space. The main window is opened off screen, in both themes, at the
# work area of the session (1920x1032 on a 1080p screen at 100 percent) and at 1024x720, and measured: sidebar buttons, inline
# buttons, single-line boxes, grid rows and headers, the width that the content uses, and the free area under and beside the
# content of every view. The 1024x720 case only proves that nothing clips and that the sidebar fits.
# The visible run prints the work area of the session (SystemParameters.WorkArea).

BeforeDiscovery {
    . "$PSScriptRoot\WindowHarness.ps1"
    $script:WindowGuard = Test-WindowTestsAllowed
    if (-not $script:WindowGuard.Allowed -and -not $env:SUITE_WINDOW_GUARD_REPORTED) { Write-Host ('Window tests skipped: {0}.' -f $script:WindowGuard.Reason) }
    $design = Get-DesignSize
    $script:Cases = @(foreach ($dark in $true, $false) {
            foreach ($size in @(@(1024, 720), @($design.W, $design.H))) {
                @{ Dark = $dark; Theme = $(if ($dark) { 'dark' } else { 'light' }); W = $size[0]; H = $size[1]; Size = ('{0}x{1}' -f $size[0], $size[1]) }
            }
        })
    $script:DesignCases = @($script:Cases | Where-Object { $_.W -ne 1024 })
    . "$PSScriptRoot\WindowConfig.ps1"
    $script:ViewCases = @($script:AppConfig.Views | ForEach-Object { @{ View = $_.Name } })
}

BeforeAll {
    . "$PSScriptRoot\TestHelpers.ps1"
    . "$PSScriptRoot\WindowHarness.ps1"
    $script:WindowGuard = Test-WindowTestsAllowed
    if (-not $script:WindowGuard.Allowed) { return }
    . "$PSScriptRoot\WindowConfig.ps1"
    $script:Design = Get-DesignSize
    $script:ShellParts = Get-ShellRuntimePart
    Add-Type -AssemblyName PresentationFramework
    $wa = [System.Windows.SystemParameters]::WorkArea
    Write-Host ('Work area of this session: {0}x{1}' -f $wa.Width, $wa.Height)
    $script:Measured = @{}
    foreach ($case in @(
            @{ Dark = $true; W = 1024; H = 720 }, @{ Dark = $true; W = $script:Design.W; H = $script:Design.H },
            @{ Dark = $false; W = 1024; H = 720 }, @{ Dark = $false; W = $script:Design.W; H = $script:Design.H })) {
        $key = '{0}|{1}x{2}' -f $(if ($case.Dark) { 'dark' } else { 'light' }), $case.W, $case.H
        $case.Config = $script:AppConfig
        $case.Parts = $script:ShellParts
        $case.TempDir = $TestDrive
        $script:Measured[$key] = Invoke-WpfSta -Arguments $case -Script {
            $dsCfg = $Arguments.Config
            $dsW = Open-TestWindow -XamlPath (Join-Path $Root $dsCfg.Xaml) -Dark $Arguments.Dark -Width $Arguments.W -Height $Arguments.H
            try {
                Initialize-ShellRuntime -Window $dsW -Parts $Arguments.Parts -LogPath (Join-Path $Arguments.TempDir 'density.log')
                Show-AppView -Window $dsW -Config $dsCfg -View $dsCfg.Views[0]
                & $global:ApplyTheme -IsDark $Arguments.Dark
                Wait-Layout $dsW
                $dsSidebar = @(foreach ($dsName in $dsCfg.SidebarButtons) {
                        $dsB = $dsW.FindName($dsName); $dsP = $dsB.TransformToAncestor($dsW).Transform([System.Windows.Point]::new(0, 0))
                        [pscustomobject]@{ Name = $dsName; Top = $dsP.Y; Bottom = $dsP.Y + $dsB.ActualHeight; Height = $dsB.ActualHeight }
                    })
                $dsViews = @(foreach ($dsView in $dsCfg.Views) {
                        Show-AppView -Window $dsW -Config $dsCfg -View $dsView
                        & $global:ApplyTheme -IsDark $Arguments.Dark
                        Wait-Layout $dsW
                        $dsHost = $dsW.FindName($dsCfg.ContentHost)
                        $dsMaxBottom = 0.0; $dsMaxRight = 0.0
                        $dsButtons = @(); $dsBoxes = @(); $dsGrids = @()
                        foreach ($dsEl in (Get-VisualDescendants $dsHost)) {
                            if ($dsEl -isnot [System.Windows.FrameworkElement] -or $dsEl.ActualWidth -le 0 -or $dsEl.ActualHeight -le 0 -or -not (Test-VisibleInTree $dsEl)) { continue }
                            $dsContent = ($dsEl -is [System.Windows.Controls.Button]) -or ($dsEl -is [System.Windows.Controls.TextBox]) -or ($dsEl -is [System.Windows.Controls.ComboBox]) -or ($dsEl -is [System.Windows.Controls.CheckBox]) -or
                            ($dsEl -is [System.Windows.Controls.RadioButton]) -or ($dsEl -is [System.Windows.Controls.DataGrid]) -or ($dsEl -is [System.Windows.Controls.ListBox]) -or
                            (($dsEl -is [System.Windows.Controls.TextBlock]) -and -not [string]::IsNullOrWhiteSpace($dsEl.Text))
                            if ($dsContent) {
                                $dsP = $dsEl.TransformToAncestor($dsHost).Transform([System.Windows.Point]::new($dsEl.ActualWidth, $dsEl.ActualHeight))
                                if ($dsP.Y -gt $dsMaxBottom) { $dsMaxBottom = $dsP.Y }
                                if ($dsP.X -gt $dsMaxRight) { $dsMaxRight = $dsP.X }
                            }
                            if ($dsEl -is [System.Windows.Controls.Button] -and $dsEl.Name) { $dsButtons += [pscustomobject]@{ Name = $dsEl.Name; Height = $dsEl.ActualHeight } }
                            if ($dsEl -is [System.Windows.Controls.TextBox] -and $dsEl.Name -and -not $dsEl.AcceptsReturn) { $dsBoxes += [pscustomobject]@{ Name = $dsEl.Name; Height = $dsEl.ActualHeight } }
                            if ($dsEl -is [System.Windows.Controls.DataGrid]) {
                                $dsHdr = $null; foreach ($dsX in (Get-VisualDescendants $dsEl)) { if ($dsX -is [System.Windows.Controls.Primitives.DataGridColumnHeadersPresenter]) { $dsHdr = $dsX; break } }
                                $dsRow = $null; foreach ($dsX in (Get-VisualDescendants $dsEl)) { if ($dsX -is [System.Windows.Controls.DataGridRow]) { $dsRow = $dsX; break } }
                                $dsGrids += [pscustomobject]@{ Name = $dsEl.Name; Width = $dsEl.ActualWidth; HeaderHeight = $(if ($dsHdr) { $dsHdr.ActualHeight } else { $null }); RowHeight = $(if ($dsRow) { $dsRow.ActualHeight } else { $null }) }
                            }
                        }
                        [pscustomobject]@{ View = $dsView.Name; Width = $dsHost.ActualWidth; Height = $dsHost.ActualHeight; FreeHeight = $dsHost.ActualHeight - $dsMaxBottom; FreeWidth = $dsHost.ActualWidth - $dsMaxRight; Buttons = $dsButtons; Boxes = $dsBoxes; Grids = $dsGrids }
                    })
                [pscustomobject]@{ Width = $dsW.ActualWidth; Height = $dsW.ActualHeight; Sidebar = $dsSidebar; Views = $dsViews }
            }
            finally { $dsW.Close() }
        }
    }
}

AfterAll {
    if ($script:WindowGuard.Allowed) { Close-WpfSession }
}

Describe 'Density <Theme> <Size>' -ForEach $script:Cases -Skip:(-not $script:WindowGuard.Allowed) {
    It 'sizes every sidebar button between 34 and 36 px and fits all of them without scrolling' {
        $m = $script:Measured['{0}|{1}' -f $Theme, $Size]
        foreach ($b in $m.Sidebar) { $b.Height | Should -BeIn 34, 35, 36 -Because $b.Name }
        ($m.Sidebar | Select-Object -Last 1).Bottom | Should -BeLessThan $m.Height
    }
}

Describe 'Density at the design size <Theme> <Size>' -ForEach $script:DesignCases -Skip:(-not $script:WindowGuard.Allowed) {
    It 'sizes the inline buttons and the single-line boxes in the <View> view' -ForEach $script:ViewCases {
        $m = $script:Measured['{0}|{1}' -f $Theme, $Size].Views | Where-Object View -eq $View
        foreach ($b in $m.Buttons | Where-Object { $_.Name -notmatch '^btnDetection$|^btnInstalled$|^btnOptions$' }) { $b.Height | Should -BeLessOrEqual 28.5 -Because $b.Name }
        foreach ($b in $m.Boxes) { $b.Height | Should -BeIn 26, 27, 28 -Because $b.Name }
    }

    It 'gives grid rows 24 to 26 px and headers 26 px in the <View> view' -ForEach $script:ViewCases {
        $m = $script:Measured['{0}|{1}' -f $Theme, $Size].Views | Where-Object View -eq $View
        foreach ($g in $m.Grids) {
            if ($null -ne $g.HeaderHeight) { $g.HeaderHeight | Should -BeLessOrEqual 26.5 -Because $g.Name }
            if ($null -ne $g.RowHeight) { [math]::Round($g.RowHeight) | Should -BeIn 24, 25, 26 -Because $g.Name }
        }
    }

    It 'lets the content fill the width in the <View> view' -ForEach $script:ViewCases {
        $mm = $script:Measured['{0}|{1}' -f $Theme, $Size]
        $m = $mm.Views | Where-Object View -eq $View
        $m.Width | Should -BeGreaterOrEqual ($mm.Width - 300)
        foreach ($g in $m.Grids) { $g.Width | Should -BeGreaterOrEqual ($m.Width * 0.3) -Because $g.Name }
    }
}

# No vacant space at the maximized work area: under the last content of a view at most 120 px are free, and right of it at most 160 px.
Describe 'Vacant space <Theme> <Size>' -ForEach $script:DesignCases -Skip:(-not $script:WindowGuard.Allowed) {
    It 'leaves no large vacant area in any view' {
        $offender = @($script:Measured['{0}|{1}' -f $Theme, $Size].Views | Where-Object { $_.FreeHeight -gt 120 -or $_.FreeWidth -gt 160 } |
                ForEach-Object { '{0}: {1:N0} px free below, {2:N0} px free right (view {3:N0}x{4:N0})' -f $_.View, $_.FreeHeight, $_.FreeWidth, $_.Width, $_.Height })
        $offender | Should -BeNullOrEmpty -Because ('these views leave a vacant area: ' + ($offender -join '; '))
    }
}

Describe 'Measured values <Theme> <Size>' -ForEach $script:DesignCases -Skip:(-not $script:WindowGuard.Allowed) {
    It 'lists the free area and the grid sizes of every view' {
        $m = $script:Measured['{0}|{1}' -f $Theme, $Size]
        foreach ($v in $m.Views) {
            Write-Host ('{0} {1}: view {2:N0}x{3:N0}, free below {4:N0}, free right {5:N0}; grids: {6}' -f $Theme, $v.View, $v.Width, $v.Height, $v.FreeHeight, $v.FreeWidth, ((@($v.Grids) | ForEach-Object { '{0} w{1:N0} header {2:N1} row {3:N1}' -f $_.Name, $_.Width, $_.HeaderHeight, $_.RowHeight }) -join '; '))
        }
        @($m.Views).Count | Should -BeGreaterThan 3
    }
}
