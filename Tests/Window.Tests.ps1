#Requires -Modules Pester
# Window tests. The main window, both module views and the Options dialog (built from XAML text in the
# script) are opened off screen on an STA thread, in both themes and at 1024x720 and 1440x900, and
# measured. The app-specific names live in WindowConfig.ps1.

BeforeDiscovery {
    . "$PSScriptRoot\WindowConfig.ps1"
    . "$PSScriptRoot\WindowHarness.ps1"
    # Window tests open a window only where the guard allows it. Everywhere else they report Skipped.
    $script:WindowGuard = Test-WindowTestsAllowed
    if (-not $script:WindowGuard.Allowed) { Write-Host ('Window tests skipped: {0}.' -f $script:WindowGuard.Reason) }
    $script:Matrix = foreach ($dark in $true, $false) {
        foreach ($size in @(@(1024, 720), @(1440, 900))) {
            @{ Dark = $dark; Theme = $(if ($dark) { 'dark' } else { 'light' }); W = $size[0]; H = $size[1]; Size = ('{0}x{1}' -f $size[0], $size[1]) }
        }
    }
    $script:ViewCases = @($script:AppConfig.Views | ForEach-Object { @{ View = $_.Name } })
    $script:ViewNames = @($script:AppConfig.Views | ForEach-Object { $_.Name }) + @('Options dialog: About', 'Options dialog: Logging')
    $script:AllViewCases = @($script:ViewNames | ForEach-Object { @{ View = $_ } })

    . "$PSScriptRoot\TestHelpers.ps1"
    $script:MainNameCases = @(Get-ShellFindNameArgument '$window' | ForEach-Object { @{ Name = $_ } })
    $script:DialogNameCases = @(Get-ShellFindNameArgument '$dlg' | ForEach-Object { @{ Name = $_ } })
    $script:ModuleNameCases = @(foreach ($loader in @(@{ Module = 'DetectionTester'; Loader = 'LoadDetectionTester' }, @{ Module = 'InstalledApps'; Loader = 'LoadInstalledApps' })) {
            foreach ($n in (Get-ShellLoaderName -Loader $loader.Loader)) { @{ Module = $loader.Module; Name = $n } }
        })
    $dialogCount = @((Get-ShellAst).FindAll({ param($n) $n -is [System.Management.Automation.Language.StringConstantExpressionAst] -and $n.Value -match '^\s*<Controls:MetroWindow' }, $true)).Count
    $script:DialogCases = @(for ($i = 0; $i -lt $dialogCount; $i++) { @{ Index = $i } })
}

BeforeAll {
    . "$PSScriptRoot\TestHelpers.ps1"
    . "$PSScriptRoot\WindowHarness.ps1"
    # No runspace and no window before the guard has said yes.
    if (-not (Test-WindowTestsAllowed).Allowed) { return }
    . "$PSScriptRoot\KnownContrastDefects.ps1"
    . "$PSScriptRoot\WindowConfig.ps1"
    $script:Root = Split-Path -Parent $PSScriptRoot
    $script:ShellPath = Join-Path $script:Root $script:AppConfig.Shell
    $script:XamlPath = Join-Path $script:Root $script:AppConfig.Xaml
    $script:ShellText = Get-Content -LiteralPath $script:ShellPath -Raw -Encoding UTF8

    $ast = Get-ShellAst
    $script:MainNames = @(Get-ShellFindNameArgument '$window')
    $script:ModuleNames = @{ DetectionTester = @(Get-ShellLoaderName -Loader 'LoadDetectionTester'); InstalledApps = @(Get-ShellLoaderName -Loader 'LoadInstalledApps') }
    $script:DialogXaml = @($ast.FindAll({
                param($n)
                $n -is [System.Management.Automation.Language.StringConstantExpressionAst] -and $n.Value -match '^\s*<Controls:MetroWindow'
            }, $true) | ForEach-Object { $_.Value })

    # The colour literals that the shell sets at start-up and on a theme change.
    $script:ShellColors = @{}
    foreach ($k in 'DarkButtonBg', 'DarkButtonBorder', 'LightButtonBg', 'LightButtonBorder', 'TitleBarBlue', 'TitleBarBlueInactive') {
        $m = [regex]::Match($script:ShellText, ('\$script:{0}\s*=\s*\[Windows\.Media\.BrushConverter\]::new\(\)\.ConvertFrom\(''(#[0-9A-Fa-f]{{6}})''\)' -f $k))
        if ($m.Success) { $script:ShellColors[$k] = $m.Groups[1].Value }
    }
    $log = [regex]::Match($script:ShellText, '\$hex\s*=\s*if\s*\(\$IsDark\)\s*\{\s*''(#[0-9A-Fa-f]{6})''\s*\}\s*else\s*\{\s*''(#[0-9A-Fa-f]{6})''\s*\}')
    if ($log.Success) { $script:ShellColors.LogDark = $log.Groups[1].Value; $script:ShellColors.LogLight = $log.Groups[2].Value }

    $script:ShellParts = Get-ShellRuntimePart
    $script:Measured = @{}
    foreach ($case in @(
            @{ Dark = $true; W = 1024; H = 720 }, @{ Dark = $true; W = 1440; H = 900 },
            @{ Dark = $false; W = 1024; H = 720 }, @{ Dark = $false; W = 1440; H = 900 })) {
        $key = '{0}|{1}x{2}' -f $(if ($case.Dark) { 'dark' } else { 'light' }), $case.W, $case.H
        $case.Config = $script:AppConfig
        $case.Colors = $script:ShellColors
        $case.Parts = $script:ShellParts
        $case.TempDir = $TestDrive
        $case.DialogXaml = $script:DialogXaml[0]
        $script:Measured[$key] = Invoke-WpfSta -Arguments $case -Script $script:MeasureAppScript
    }
}

AfterAll { Close-WpfSession }

Describe 'Main window XAML' -Skip:(-not $script:WindowGuard.Allowed) {
    It 'loads as a MetroWindow with the app title and a minimum size that fits 1024x720' {
        $r = Invoke-WpfSta -Arguments @{ Xaml = $script:AppConfig.Xaml } -Script {
            $w = Open-TestWindow -XamlPath (Join-Path $Root $Arguments.Xaml)
            $o = [PSCustomObject]@{ Type = $w.GetType().FullName; Title = $w.Title; MinW = $w.MinWidth; MinH = $w.MinHeight }
            $w.Close(); $o
        }
        $r.Type | Should -Be 'MahApps.Metro.Controls.MetroWindow'
        $r.Title | Should -Be $script:AppConfig.Title
        $r.MinW | Should -BeLessOrEqual 1024
        $r.MinH | Should -BeLessOrEqual 720
    }

    It 'declares each x:Name once' {
        $names = [regex]::Matches((Get-Content -LiteralPath $script:XamlPath -Raw), 'x:Name="([^"]+)"') | ForEach-Object { $_.Groups[1].Value }
        ($names | Group-Object | Where-Object Count -gt 1 | ForEach-Object Name) | Should -BeNullOrEmpty
    }

    It 'finds the script look-ups of the window' {
        $script:MainNames.Count | Should -BeGreaterThan 8
    }

    It 'has the control <Name> that the script looks up' -ForEach $script:MainNameCases {
        (Get-Content -LiteralPath $script:XamlPath -Raw) | Should -Match ('x:Name="{0}"' -f [regex]::Escape($Name))
    }

    It 'has the sidebar buttons and the content host that the test configuration names' {
        $xaml = Get-Content -LiteralPath $script:XamlPath -Raw
        foreach ($n in $script:AppConfig.SidebarButtons + $script:AppConfig.ContentHost + 'txtModuleTitle' + 'txtModuleSubtitle' + 'txtStatus' + $script:AppConfig.LogLabel) {
            $xaml | Should -Match ('x:Name="{0}"' -f $n)
        }
    }

    It 'sets the texts that the test configuration shows for the <Module> module' -ForEach @(@{ Module = 'DetectionTester' }, @{ Module = 'InstalledApps' }) {
        $m = $script:AppConfig.Modules[$Module]
        foreach ($text in $m.Title, $m.Subtitle, $m.Status) {
            $script:ShellText.Contains("'$text'") | Should -BeTrue -Because "the shell sets '$text'"
        }
    }

    It 'finds the theme colours in the shell' {
        foreach ($k in 'DarkButtonBg', 'DarkButtonBorder', 'LightButtonBg', 'LightButtonBorder', 'TitleBarBlue', 'TitleBarBlueInactive', 'LogDark', 'LogLight') {
            $script:ShellColors[$k] | Should -Match '^#[0-9A-Fa-f]{6}$' -Because "the shell has a colour for $k"
        }
    }
}

Describe 'Module XAML' -Skip:(-not $script:WindowGuard.Allowed) {
    It 'has the control <Name> that the <Module> loader looks up' -ForEach $script:ModuleNameCases {
        $xaml = Get-Content -LiteralPath (Join-Path $script:Root $script:AppConfig.Modules[$Module].Xaml) -Raw
        $xaml | Should -Match ('x:Name="{0}"' -f [regex]::Escape($Name))
    }

    It 'declares each x:Name of the <Module> module once' -ForEach @(@{ Module = 'DetectionTester' }, @{ Module = 'InstalledApps' }) {
        $names = [regex]::Matches((Get-Content -LiteralPath (Join-Path $script:Root $script:AppConfig.Modules[$Module].Xaml) -Raw), 'x:Name="([^"]+)"') | ForEach-Object { $_.Groups[1].Value }
        ($names | Group-Object -CaseSensitive | Where-Object Count -gt 1 | ForEach-Object Name) | Should -BeNullOrEmpty
    }

    It 'declares no two names of the <Module> module that differ only in case' -ForEach @(@{ Module = 'DetectionTester' }, @{ Module = 'InstalledApps' }) {
        $names = [regex]::Matches((Get-Content -LiteralPath (Join-Path $script:Root $script:AppConfig.Modules[$Module].Xaml) -Raw), 'x:Name="([^"]+)"') | ForEach-Object { $_.Groups[1].Value }
        ($names | Group-Object | Where-Object Count -gt 1 | ForEach-Object Name) | Should -BeNullOrEmpty
    }

    It 'loads the <Module> module root and finds all its controls' -ForEach @(@{ Module = 'DetectionTester' }, @{ Module = 'InstalledApps' }) {
        $cases = @($script:ModuleNames[$Module])
        $r = Invoke-WpfSta -Arguments @{ Xaml = $script:AppConfig.Modules[$Module].Xaml; Names = $cases } -Script {
            [xml]$x = Get-Content -LiteralPath (Join-Path $Root $Arguments.Xaml) -Raw -Encoding UTF8
            $moduleRoot = [System.Windows.Markup.XamlReader]::Load((New-Object System.Xml.XmlNodeReader $x))
            [PSCustomObject]@{ Missing = @($Arguments.Names | Where-Object { -not $moduleRoot.FindName($_) }); Count = @($Arguments.Names).Count }
        }
        $r.Missing | Should -BeNullOrEmpty
        $r.Count | Should -BeGreaterThan 5
    }

    It 'shows the first input panel of the Detection Tester and hides the others' {
        $r = Invoke-WpfSta -Arguments @{ Xaml = $script:AppConfig.Modules.DetectionTester.Xaml; Panels = $script:AppConfig.Panels } -Script {
            [xml]$x = Get-Content -LiteralPath (Join-Path $Root $Arguments.Xaml) -Raw -Encoding UTF8
            $moduleRoot = [System.Windows.Markup.XamlReader]::Load((New-Object System.Xml.XmlNodeReader $x))
            $o = foreach ($p in $Arguments.Panels) { [PSCustomObject]@{ Name = $p; Vis = [string]$moduleRoot.FindName($p).Visibility } }
            , @($o)
        }
        ($r | Where-Object Name -eq $script:AppConfig.Panels[0]).Vis | Should -Be 'Visible'
        @($r | Where-Object { $_.Name -ne $script:AppConfig.Panels[0] -and $_.Vis -ne 'Collapsed' }) | Should -BeNullOrEmpty
    }

    It 'lists the five detection types in the order the shell expects' {
        $r = Invoke-WpfSta -Arguments @{ Xaml = $script:AppConfig.Modules.DetectionTester.Xaml } -Script {
            [xml]$x = Get-Content -LiteralPath (Join-Path $Root $Arguments.Xaml) -Raw -Encoding UTF8
            $moduleRoot = [System.Windows.Markup.XamlReader]::Load((New-Object System.Xml.XmlNodeReader $x))
            , @($moduleRoot.FindName('cboType').Items | ForEach-Object { [string]$_.Content })
        }
        ($r -join '|') | Should -Be 'RegistryKeyValue|RegistryKey|File|Script|Compound (import-only)'
    }

    It 'lists the operators that the shell sends to the module: <Combo>' -ForEach @(
        @{ Combo = 'rkvOperator'; Expected = 'IsEquals|GreaterEquals|GreaterThan|LessEquals|LessThan' }
        @{ Combo = 'filePropType'; Expected = 'Existence|Version' }
        @{ Combo = 'fileOperator'; Expected = 'GreaterEquals|GreaterThan|IsEquals|LessEquals|LessThan' }
    ) {
        $r = Invoke-WpfSta -Arguments @{ Xaml = $script:AppConfig.Modules.DetectionTester.Xaml; Combo = $Combo } -Script {
            [xml]$x = Get-Content -LiteralPath (Join-Path $Root $Arguments.Xaml) -Raw -Encoding UTF8
            $moduleRoot = [System.Windows.Markup.XamlReader]::Load((New-Object System.Xml.XmlNodeReader $x))
            , @($moduleRoot.FindName($Arguments.Combo).Items | ForEach-Object { [string]$_.Content })
        }
        ($r -join '|') | Should -Be $Expected
    }
}

Describe 'Dialog windows' -Skip:(-not $script:WindowGuard.Allowed) {
    It 'finds the dialog XAML in the script' {
        $script:DialogXaml.Count | Should -BeGreaterThan 0
    }

    It 'parses dialog <Index> as a MetroWindow with named controls' -ForEach $script:DialogCases {
        $r = Invoke-WpfSta -Arguments @{ Xaml = $script:DialogXaml[$Index] } -Script {
            [xml]$x = $Arguments.Xaml
            $w = [System.Windows.Markup.XamlReader]::Load((New-Object System.Xml.XmlNodeReader $x))
            $names = [regex]::Matches($Arguments.Xaml, 'x:Name="([^"]+)"') | ForEach-Object { $_.Groups[1].Value }
            [PSCustomObject]@{ Type = $w.GetType().Name; Missing = @($names | Where-Object { -not $w.FindName($_) }); Count = @($names).Count }
        }
        $r.Type | Should -Be 'MetroWindow'
        $r.Missing | Should -BeNullOrEmpty
        $r.Count | Should -BeGreaterThan 0
    }

    It 'has the control <Name> that the script looks up in the dialog' -ForEach $script:DialogNameCases {
        $script:DialogXaml[0] | Should -Match ('x:Name="{0}"' -f [regex]::Escape($Name))
    }

    It 'gives every dialog a title that differs from the main window title' {
        $titles = @([regex]::Matches($script:ShellText, "-Title\s+'([^']+)'") | ForEach-Object { $_.Groups[1].Value })
        $titles += @($script:DialogXaml | ForEach-Object { if ($_ -match '\sTitle="([^"]*)"') { $Matches[1] } } | Where-Object { $_ })
        $titles.Count | Should -BeGreaterThan 0
        $titles | Should -Not -Contain $script:AppConfig.Title
        foreach ($t in $titles) { $t.Trim() | Should -Not -BeNullOrEmpty }
    }

    It 'gives the dialogs in XAML different titles' {
        $titles = @($script:DialogXaml | ForEach-Object { if ($_ -match '\sTitle="([^"]*)"') { $Matches[1] } } | Where-Object { $_ })
        ($titles | Select-Object -Unique).Count | Should -Be $titles.Count
    }

    # The theme default title bar of the light theme has 3.3:1 against the title text. The main window sets its own
    # brushes; the dialog must set the same ones, and the harness applies them to the dialog it measures.
    It 'gives the Options dialog the title bar brushes of the main window in the light theme' {
        $text = Get-ShellBlockText -Name 'ShowOptionsDialog'
        $text | Should -Match '\$dlg\.WindowTitleBrush\s*=\s*\$script:TitleBarBlue\b'
        $text | Should -Match '\$dlg\.NonActiveWindowTitleBrush\s*=\s*\$script:TitleBarBlueInactive\b'
    }

    It 'shows the dialog title "Options" on the title bar' {
        $r = Invoke-WpfSta -Arguments @{ Xaml = $script:DialogXaml[0] } -Script {
            $d = Open-TestWindow -XamlText $Arguments.Xaml -Width 640 -Height 440
            $o = $d.Title; $d.Close(); $o
        }
        $r | Should -Be 'Options'
    }
}

Describe 'Layout <Theme> <Size>' -ForEach $script:Matrix -Skip:(-not $script:WindowGuard.Allowed) {
    It 'has no clipped text in the <View> view' -ForEach $script:AllViewCases {
        $m = ($script:Measured['{0}|{1}' -f $Theme, $Size] | Where-Object View -eq $View)
        $m | Should -Not -BeNullOrEmpty
        $m.Clipped | ForEach-Object { '{0}: needs {1}px, has {2}px' -f $_.Text, $_.Needed, $_.Actual } | Should -BeNullOrEmpty
    }

    It 'keeps every button, box and list inside the window in the <View> view' -ForEach $script:AllViewCases {
        $m = ($script:Measured['{0}|{1}' -f $Theme, $Size] | Where-Object View -eq $View)
        $m.Offscreen | ForEach-Object { '{0} at {1},{2} size {3}x{4}' -f $_.Name, $_.X, $_.Y, $_.W, $_.H } | Should -BeNullOrEmpty
    }
}

Describe 'Contrast <Theme>' -ForEach @(@{ Theme = 'dark' }, @{ Theme = 'light' }) -Skip:(-not $script:WindowGuard.Allowed) {
    It 'shows every text at 4.5:1 (3:1 for large text) in the <View> view, except the listed known defects' -ForEach $script:AllViewCases {
        $m = ($script:Measured['{0}|1440x900' -f $Theme] | Where-Object View -eq $View)
        $m.Text.Count | Should -BeGreaterThan 3
        $known = @(Get-KnownContrastDefect -Theme $Theme)
        $bad = $m.Text | Where-Object { $min = if (Test-LargeText $_) { 3.0 } else { 4.5 }; $_.Ratio -lt $min -and ($known -notcontains ('text:{0}' -f $_.Text)) }
        $bad | ForEach-Object { '{0} "{1}" {2} on {3} = {4:N2}' -f $_.Name, $_.Text, $_.Foreground, $_.Backdrop, $_.Ratio } | Should -BeNullOrEmpty
    }

    It 'keeps the title text readable on the active title bar' {
        $tb = ($script:Measured['{0}|1440x900' -f $Theme] | Select-Object -First 1).TitleBar
        $tb.Active | Should -BeGreaterOrEqual 4.5
    }

    It 'meets the standard on every control, except the listed known defects' {
        $known = Get-KnownContrastDefect -Theme $Theme
        $found = @(Get-ContrastViolationKey -Views @($script:Measured['{0}|1440x900' -f $Theme]))
        $new = @($found | Where-Object { $known -notcontains $_ })
        $fixed = @($known | Where-Object { $found -notcontains $_ })
        $new | Should -BeNullOrEmpty -Because 'a new contrast violation appeared'
        $fixed | Should -BeNullOrEmpty -Because 'a known defect is fixed: remove it from Get-KnownContrastDefect'
    }
}

# The handlers of the start script run here on the real controls of a real window. The shell's own script blocks
# (loaders, click handlers, theme code) are cut out of the script and defined in the same STA runspace, and the tests
# click the buttons and read the result back. Only the dialogs of the operating system are left out, because they
# block the thread. The logic of the handlers with plain objects is in Shell.Logic.Tests.ps1. All window tests are in
# this one file, because a second STA runspace in the same process hangs the theme manager.
Describe 'Handlers on the real controls' -Skip:(-not $script:WindowGuard.Allowed) {
    BeforeAll {
        $null = Invoke-WpfSta -Arguments @{ Parts = $script:ShellParts; TempDir = $TestDrive } -Script {
            $global:HW = Open-TestWindow -XamlPath (Join-Path $Root 'MainWindow.xaml')
            Initialize-ShellRuntime -Window $global:HW -Parts $Arguments.Parts -LogPath (Join-Path $Arguments.TempDir 'handlers.log')
            & $global:ShowModule 'DetectionTester'
        }
        $script:Dir = Join-Path $TestDrive 'handler files'
        New-Item -ItemType Directory -Path $script:Dir -Force | Out-Null
        Set-Content -LiteralPath (Join-Path $script:Dir 'a.txt') -Value 'x'
    }
    AfterAll {
        $null = Invoke-WpfSta -Script { if ($global:HW) { $global:HW.Close() } }
    }

    Context 'Detection Tester handlers' {
        It 'shows only the input panel of <Type>' -ForEach @(
            @{ Index = 0; Type = 'RegistryKeyValue'; Panel = 'panelRegistryKeyValue' }
            @{ Index = 1; Type = 'RegistryKey'; Panel = 'panelRegistryKey' }
            @{ Index = 2; Type = 'File'; Panel = 'panelFile' }
            @{ Index = 3; Type = 'Script'; Panel = 'panelScript' }
            @{ Index = 4; Type = 'Compound (import-only)'; Panel = 'panelCompound' }
        ) {
            $visible = Invoke-WpfSta -Arguments @{ Index = $Index } -Script {
                $c = $global:DTContext
                $c.cboType.SelectedIndex = $Arguments.Index
                , @('panelRegistryKeyValue', 'panelRegistryKey', 'panelFile', 'panelScript', 'panelCompound' | Where-Object { [string]$c[$_].Visibility -eq 'Visible' })
            }
            @($visible) | Should -Be @($Panel)
        }

        It 'runs a file existence test from the window, in positive and in negative mode' {
            $r = Invoke-WpfSta -Arguments @{ Dir = $script:Dir } -Script {
                $c = $global:DTContext
                Invoke-Click $c.btnClearHistory
                $c.chkNegativeMode.IsChecked = $false
                $c.cboType.SelectedIndex = 2
                $c.fileFilePath.Text = $Arguments.Dir; $c.fileFileName.Text = 'a.txt'; $c.filePropType.SelectedIndex = 0
                Invoke-Click $c.btnTest
                $first = [PSCustomObject]@{ Badge = $c.txtResultBadge.Text; Details = $c.txtResultDetails.Text; Rows = $c.dgHistory.Items.Count }
                $c.chkNegativeMode.IsChecked = $true
                Invoke-Click $c.btnTest
                $c.chkNegativeMode.IsChecked = $false
                $second = [PSCustomObject]@{ Badge = $c.txtResultBadge.Text; Rows = $c.dgHistory.Items.Count; Newest = $c.dgHistory.Items[0] | Select-Object Mode, Result, Type, Target }
                , @($first, $second)
            }
            $r[0].Badge | Should -Be 'Pass'
            $r[0].Details | Should -BeLike 'File exists: *'
            $r[0].Rows | Should -Be 1
            $r[1].Badge | Should -Be 'Fail'
            $r[1].Rows | Should -Be 2
            $r[1].Newest.Mode | Should -Be 'Negative'
            $r[1].Newest.Result | Should -Be 'Fail'
            $r[1].Newest.Type | Should -Be 'File'
            $r[1].Newest.Target | Should -Be (Join-Path $script:Dir 'a.txt')
        }

        It 'runs a registry value test from the window against a throwaway key' {
            $key = New-TestRegistryKey -Value @{ Version = '1.2.3' }
            try {
                $r = Invoke-WpfSta -Arguments @{ Key = $key } -Script {
                    $c = $global:DTContext
                    Invoke-Click $c.btnClearHistory
                    $c.cboType.SelectedIndex = 0
                    $c.rkvHiveHKCU.IsChecked = $true
                    $c.rkvKeyPath.Text = $Arguments.Key; $c.rkvValueName.Text = 'Version'; $c.rkvExpected.Text = '1.2.3'; $c.rkvOperator.SelectedIndex = 1
                    Invoke-Click $c.btnTest
                    $pass = $c.txtResultBadge.Text
                    $c.rkvExpected.Text = '1.10'
                    Invoke-Click $c.btnTest
                    $stillPass = $c.txtResultBadge.Text
                    $c.rkvExpected.Text = '2.0'
                    Invoke-Click $c.btnTest
                    [PSCustomObject]@{ Pass = $pass; StillPass = $stillPass; Fail = $c.txtResultBadge.Text; Rows = $c.dgHistory.Items.Count }
                }
            }
            finally { Remove-TestRegistryKey -Relative $key }
            $r.Pass | Should -Be 'Pass'
            $r.StillPass | Should -Be 'Fail'
            $r.Fail | Should -Be 'Fail'
            $r.Rows | Should -Be 3
        }

        It 'says "Invalid clause" and keeps the history when the key path is empty' {
            $r = Invoke-WpfSta -Script {
                $c = $global:DTContext
                Invoke-Click $c.btnClearHistory
                $c.cboType.SelectedIndex = 1
                $c.rkKeyPath.Text = ''
                Invoke-Click $c.btnTest
                [PSCustomObject]@{ Badge = $c.txtResultBadge.Text; Details = $c.txtResultDetails.Text; Rows = $c.dgHistory.Items.Count }
            }
            $r.Badge | Should -Be 'Invalid clause'
            $r.Details | Should -Be 'Key path is required.'
            $r.Rows | Should -Be 0
        }

        It 'runs a script test in the background and keeps the window responsive: <Case>' -ForEach @(
            @{ Case = 'output means Pass'; Script = '"found"'; Negative = $false; Badge = 'Pass'; Mode = 'Positive' }
            @{ Case = 'no output means Fail'; Script = '$null'; Negative = $false; Badge = 'Fail'; Mode = 'Positive' }
            @{ Case = 'output in negative mode means Fail'; Script = '"found"'; Negative = $true; Badge = 'Fail'; Mode = 'Negative' }
        ) {
            $r = Invoke-WpfSta -Arguments @{ Script = $Script; Negative = $Negative } -Script {
                $c = $global:DTContext
                Invoke-Click $c.btnClearHistory
                $c.cboType.SelectedIndex = 3
                $c.scriptText.Text = $Arguments.Script
                $c.chkNegativeMode.IsChecked = $Arguments.Negative
                Invoke-Click $c.btnTest
                $during = [PSCustomObject]@{ Enabled = $c.btnTest.IsEnabled; Badge = $c.txtResultBadge.Text }
                $done = Wait-Until -Window $global:HW -TimeoutSeconds 60 -Condition { $global:DTContext.btnTest.IsEnabled }
                $c.chkNegativeMode.IsChecked = $false
                [PSCustomObject]@{ During = $during; Done = $done; Badge = $c.txtResultBadge.Text; Rows = $c.dgHistory.Items.Count; Row = ($c.dgHistory.Items | Select-Object -First 1 | Select-Object Type, Mode, Result, Found) }
            }
            $r.During.Enabled | Should -BeFalse
            $r.During.Badge | Should -Be 'Running...'
            $r.Done | Should -BeTrue
            $r.Badge | Should -Be $Badge
            $r.Rows | Should -Be 1
            $r.Row.Type | Should -Be 'Script'
            $r.Row.Mode | Should -Be $Mode
            $r.Row.Result | Should -Be $Badge
        }

        It 'runs a compound test from the window with the imported clauses' {
            $r = Invoke-WpfSta -Arguments @{ Dir = $script:Dir } -Script {
                $c = $global:DTContext
                Invoke-Click $c.btnClearHistory
                $c.cboType.SelectedIndex = 4
                $c.compoundConnector.Text = 'And'
                $c.CompoundClauseData = @(
                    @{ Type = 'RegistryKey'; RegistryKeyRelative = 'Environment'; Hive = 'CurrentUser'; Is64Bit = $true }
                    @{ Type = 'File'; FilePath = $Arguments.Dir; FileName = 'a.txt'; PropertyType = 'Existence' }
                )
                Invoke-Click $c.btnTest
                $pass = [PSCustomObject]@{ Badge = $c.txtResultBadge.Text; Found = $c.dgHistory.Items[0].Found }
                $c.CompoundClauseData = @(@{ Type = 'RegistryKey'; RegistryKeyRelative = 'Environment'; Hive = 'CurrentUser' }, @{ Type = 'File'; FilePath = $Arguments.Dir; FileName = 'missing.txt' })
                Invoke-Click $c.btnTest
                , @($pass, [PSCustomObject]@{ Badge = $c.txtResultBadge.Text; Found = $c.dgHistory.Items[0].Found })
            }
            $r[0].Badge | Should -Be 'Pass'
            $r[0].Found | Should -Be '2/2 passed'
            $r[1].Badge | Should -Be 'Fail'
            $r[1].Found | Should -Be '1/2 passed'
        }

        It 'shows "Use selected" only for a selected history row and restores the clause from it' {
            $r = Invoke-WpfSta -Arguments @{ Dir = $script:Dir } -Script {
                $c = $global:DTContext
                Invoke-Click $c.btnClearHistory
                $c.cboType.SelectedIndex = 2
                $c.fileFilePath.Text = $Arguments.Dir; $c.fileFileName.Text = 'a.txt'; $c.filePropType.SelectedIndex = 0
                $c.chkNegativeMode.IsChecked = $true
                Invoke-Click $c.btnTest
                $c.chkNegativeMode.IsChecked = $false
                $c.fileFileName.Text = 'changed'; $c.cboType.SelectedIndex = 0
                $before = [string]$c.btnRepopulate.Visibility
                $c.dgHistory.SelectedIndex = 0
                $selected = [string]$c.btnRepopulate.Visibility
                Invoke-Click $c.btnRepopulate
                $restored = [PSCustomObject]@{ Type = $c.cboType.SelectedIndex; Name = $c.fileFileName.Text; Negative = [bool]$c.chkNegativeMode.IsChecked }
                $c.dgHistory.SelectedIndex = -1
                $c.chkNegativeMode.IsChecked = $false
                [PSCustomObject]@{ Before = $before; Selected = $selected; Restored = $restored; After = [string]$c.btnRepopulate.Visibility }
            }
            $r.Before | Should -Be 'Collapsed'
            $r.Selected | Should -Be 'Visible'
            $r.Restored.Type | Should -Be 2
            $r.Restored.Name | Should -Be 'a.txt'
            $r.Restored.Negative | Should -BeTrue
            $r.After | Should -Be 'Collapsed'
        }

        It 'empties the history with the Clear button' {
            $r = Invoke-WpfSta -Arguments @{ Dir = $script:Dir } -Script {
                $c = $global:DTContext
                $c.cboType.SelectedIndex = 2
                $c.fileFilePath.Text = $Arguments.Dir; $c.fileFileName.Text = 'a.txt'; $c.filePropType.SelectedIndex = 0
                Invoke-Click $c.btnTest
                $before = $c.dgHistory.Items.Count
                Invoke-Click $c.btnClearHistory
                [PSCustomObject]@{ Before = $before; After = $c.dgHistory.Items.Count }
            }
            $r.Before | Should -BeGreaterThan 0
            $r.After | Should -Be 0
        }
    }

    Context 'Installed Applications handlers' {
        It 'lists the installed applications, filters them by text and by scope, and sends one to the Detection Tester' {
            $key = New-TestRegistryKey -Parent 'Software\Microsoft\Windows\CurrentVersion\Uninstall' -Value @{ DisplayName = 'ZZTEST-DT Handler App'; DisplayVersion = '7.8.9'; Publisher = 'ZZ' }
            try {
                $r = Invoke-WpfSta -Script {
                    & $global:ShowModule 'InstalledApps'
                    $i = $global:IAContext
                    Invoke-Click $i.btnRefresh
                    $all = $i.dgApps.Items.Count
                    $i.txtFilter.Text = 'ZZTEST-DT Handler'
                    $byText = $i.dgApps.Items.Count
                    $i.cboScope.SelectedIndex = 1
                    $machine = $i.dgApps.Items.Count
                    $i.cboScope.SelectedIndex = 2
                    $user = $i.dgApps.Items.Count
                    $i.cboScope.SelectedIndex = 0
                    $i.txtFilter.Text = 'zzzz-no-such-application-zzzz'
                    $none = $i.dgApps.Items.Count
                    $i.txtFilter.Text = 'ZZTEST-DT Handler'
                    $buttonsBefore = [string]$i.btnUseForDetection.Visibility
                    $i.dgApps.SelectedIndex = 0
                    $buttonsAfter = [string]$i.btnUseForDetection.Visibility
                    $detail = [PSCustomObject]@{ Name = $i.dtlDisplayName.Text; Version = $i.dtlDisplayVersion.Text; ArchScope = $i.dtlArchScope.Text }
                    Invoke-Click $i.btnUseForDetection
                    $d = $global:DTContext
                    $content = [object]::ReferenceEquals($global:contentHost.Content, $global:DTRoot)
                    [PSCustomObject]@{
                        All = $all; ByText = $byText; Machine = $machine; User = $user; None = $none; ButtonsBefore = $buttonsBefore; ButtonsAfter = $buttonsAfter; Detail = $detail
                        Sent = [PSCustomObject]@{ Content = $content; Type = $d.cboType.SelectedIndex; Path = $d.rkvKeyPath.Text; Name = $d.rkvValueName.Text; Expected = $d.rkvExpected.Text; HKCU = [bool]$d.rkvHiveHKCU.IsChecked; Title = $global:txtModuleTitle.Text }
                    }
                }
            }
            finally { Remove-TestRegistryKey -Relative $key }
            $r.All | Should -BeGreaterThan 0
            $r.ByText | Should -Be 1
            $r.Machine | Should -Be 0 -Because 'the throwaway entry is a per-user entry'
            $r.User | Should -Be 1
            $r.None | Should -Be 0
            $r.ButtonsBefore | Should -Be 'Collapsed'
            $r.ButtonsAfter | Should -Be 'Visible'
            $r.Detail.Name | Should -Be 'ZZTEST-DT Handler App'
            $r.Detail.Version | Should -Be '7.8.9'
            $r.Detail.ArchScope | Should -Be 'x64 / User'
            $r.Sent.Content | Should -BeTrue
            $r.Sent.Type | Should -Be 0
            $r.Sent.Path | Should -Be ('SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\' + (Split-Path -Path $key -Leaf))
            $r.Sent.Name | Should -Be 'DisplayVersion'
            $r.Sent.Expected | Should -Be '7.8.9'
            $r.Sent.HKCU | Should -BeTrue
            $r.Sent.Title | Should -Be 'Detection Tester'
        }
    }

    Context 'Shell handlers' {
        It 'shows the title, the subtitle and the status of each module and remembers the active module' {
            $r = Invoke-WpfSta -Script {
                $o = foreach ($m in 'InstalledApps', 'DetectionTester') {
                    & $global:ShowModule $m
                    [PSCustomObject]@{ Module = $m; Title = $global:txtModuleTitle.Text; Subtitle = $global:txtModuleSubtitle.Text; Status = $global:txtStatus.Text; Active = $global:Prefs.ActiveModule }
                }
                , @($o)
            }
            $r[0].Title | Should -Be 'Installed Applications'
            $r[0].Status | Should -Be 'Installed Applications selected.'
            $r[0].Active | Should -Be 'InstalledApps'
            $r[1].Title | Should -Be 'Detection Tester'
            $r[1].Active | Should -Be 'DetectionTester'
            $r[1].Subtitle | Should -Not -BeNullOrEmpty
        }

        It 'applies the <Name> theme with the theme name, the label, the log colour and the sidebar colours' -ForEach @(
            @{ Name = 'light'; Dark = $false; Theme = 'Light.Blue'; Label = 'Light Theme'; Log = '#FF595959'; Button = '#FF0078D4'; Border = '#FF006CBE'; Title = '#FF0078D4' }
            @{ Name = 'dark'; Dark = $true; Theme = 'Dark.Steel'; Label = 'Dark Theme'; Log = '#FFB0B0B0'; Button = '#FF1E1E1E'; Border = '#FF555555'; Title = $null }
        ) {
            $r = Invoke-WpfSta -Arguments @{ Dark = $Dark } -Script {
                & $global:ApplyTheme -IsDark $Arguments.Dark
                Wait-Layout $global:HW
                [PSCustomObject]@{
                    Theme = [ControlzEx.Theming.ThemeManager]::Current.DetectTheme($global:HW).Name
                    Label = $global:txtThemeLabel.Text
                    Log = $global:lblLogOutput.Foreground.Color.ToString()
                    Button = $global:HW.FindName('btnDetection').Background.Color.ToString()
                    Border = $global:HW.FindName('btnOptions').BorderBrush.Color.ToString()
                    TitleLocal = $global:HW.ReadLocalValue([MahApps.Metro.Controls.MetroWindow]::WindowTitleBrushProperty) -is [System.Windows.Media.SolidColorBrush]
                    Title = $(if ($global:HW.WindowTitleBrush -is [System.Windows.Media.SolidColorBrush]) { $global:HW.WindowTitleBrush.Color.ToString() })
                    Pref = $global:Prefs.DarkMode
                }
            }
            $r.Theme | Should -Be $Theme
            $r.Label | Should -Be $Label
            $r.Log | Should -Be $Log
            $r.Button | Should -Be $Button
            $r.Border | Should -Be $Border
            $r.Pref | Should -Be $Dark
            if ($Dark) { $r.TitleLocal | Should -BeFalse } else { $r.Title | Should -Be $Title }
        }
    }
}
