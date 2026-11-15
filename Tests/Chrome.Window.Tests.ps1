#Requires -Modules @{ ModuleName = 'Pester'; ModuleVersion = '5.1.0' }
# Window test for the chrome and the buttons of the main window and of every dialog, in both themes.
# Every dialog is built from its XAML text, given the theme the way the script gives it (Set-DialogTheme, and the inactive title bar
# brush of the main window), and shown. Rules:
#   - the dialog title bar brush (active and inactive) is the title bar brush of the main window
#   - the title bar text of the dialog is readable (4.5:1) on that brush
#   - every button has bold text (SemiBold or heavier, 600; a normal weight is a defect)
#   - no button has a border thicker than 1 px, and no button border has the text color of the theme
# The Open and Save dialogs of Windows (Microsoft.Win32 file dialogs) are drawn by the operating system, not by the app. They are
# outside this test, and a test below counts them so that a new app-drawn dialog cannot be missed.

BeforeDiscovery {
    . "$PSScriptRoot\WindowHarness.ps1"
    $script:WindowGuard = Test-WindowTestsAllowed
    if (-not $script:WindowGuard.Allowed -and -not $env:SUITE_WINDOW_GUARD_REPORTED) { Write-Host ('Window tests skipped: {0}.' -f $script:WindowGuard.Reason) }
    $script:ThemeCases = @(@{ Dark = $true; Theme = 'dark' }, @{ Dark = $false; Theme = 'light' })
}

BeforeAll {
    . "$PSScriptRoot\TestHelpers.ps1"
    . "$PSScriptRoot\WindowHarness.ps1"
    $script:WindowGuard = Test-WindowTestsAllowed
    if (-not $script:WindowGuard.Allowed) { return }
    . "$PSScriptRoot\WindowConfig.ps1"
    $script:RepoDir = Split-Path -Parent $PSScriptRoot
    $script:ShellText = Get-Content -LiteralPath (Join-Path $script:RepoDir $script:AppConfig.Shell) -Raw -Encoding UTF8
    $dialogXaml = @((Get-ShellAst).FindAll({
                param($n)
                $n -is [System.Management.Automation.Language.StringConstantExpressionAst] -and $n.Value -match '^\s*<Controls:MetroWindow'
            }, $true) | ForEach-Object { $_.Value })
    $script:DialogXamlCount = $dialogXaml.Count
    $script:ShellParts = Get-ShellRuntimePart
    $script:Measured = @{}
    foreach ($dark in $true, $false) {
        $key = if ($dark) { 'dark' } else { 'light' }
        $script:Measured[$key] = Invoke-WpfSta -Arguments @{ Dark = $dark; Config = $script:AppConfig; Parts = $script:ShellParts; TempDir = $TestDrive; Dialogs = $dialogXaml } -Script {
            $cfg = $Arguments.Config
            $main = Open-TestWindow -XamlPath (Join-Path $Root $cfg.Xaml) -Dark $Arguments.Dark
            Initialize-ShellRuntime -Window $main -Parts $Arguments.Parts -LogPath (Join-Path $Arguments.TempDir 'chrome.log')
            Show-AppView -Window $main -Config $cfg -View $cfg.Views[0]
            & $global:ApplyTheme -IsDark $Arguments.Dark
            Wait-Layout $main
            $mainChrome = Get-SuiteWindowChrome -Window $main -Label 'main'
            $out = [pscustomobject]@{
                ThemeText = $mainChrome.ThemeText
                MainTitle = $mainChrome.Active; MainTitleInactive = $mainChrome.Inactive
                Main = @($mainChrome.Buttons); Dialogs = @()
            }
            $i = 0
            foreach ($xaml in $Arguments.Dialogs) {
                $i++
                $d = Open-TestWindow -XamlText $xaml -Dark $Arguments.Dark -Width 640 -Height 440
                $d.Resources['MahApps.Brushes.WindowTitle.NonActive'] = $global:TitleBarInactiveBrush
                Set-DialogTheme -Dialog $d -IsDark $Arguments.Dark
                Wait-Layout $d
                $c = Get-SuiteWindowChrome -Window $d -Label ([string]$d.Title)
                $out.Dialogs += [pscustomobject]@{ Title = $c.Title; Index = $i; Active = $c.Active; Inactive = $c.Inactive; TitleRatio = $c.TitleRatio; Buttons = @($c.Buttons) }
                $d.Close()
            }
            $main.Close()
            $out
        }
    }
}

AfterAll {
    if ($script:WindowGuard.Allowed) { Close-WpfSession }
}

Describe 'Chrome <Theme>' -ForEach $script:ThemeCases -Skip:(-not $script:WindowGuard.Allowed) {
    BeforeAll { $script:M = $script:Measured[$Theme] }

    It 'finds every dialog of the script' {
        # An app-drawn dialog is a ShowDialog call that is neither a Windows file dialog nor the main window.
        $shown = [regex]::Matches($script:ShellText, '\.ShowDialog\(').Count
        $native = [regex]::Matches($script:ShellText, 'New-Object\s+Microsoft\.Win32\.\w*FileDialog').Count
        $mainShown = [regex]::Matches($script:ShellText, '\$window\.ShowDialog\(\)').Count
        $themed = [regex]::Matches($script:ShellText, '(?m)^\s*(?:\[void\]\s*)?(?:\$\w+\s*=\s*)?(Show-ThemedMessage|Show-ConfirmDialog|Show-CollectionPickerDialog)\b').Count
        $expected = $shown - $native - $mainShown + $themed
        $expected | Should -BeGreaterThan 0
        $script:DialogXamlCount | Should -Be $expected -Because 'every dialog of the script needs its XAML measured here'
        @($script:M.Dialogs).Count | Should -Be $expected
    }

    It 'gives every dialog the title bar brush of the main window' {
        $bad = @($script:M.Dialogs | Where-Object { $_.Active -ne $script:M.MainTitle -or $_.Inactive -ne $script:M.MainTitleInactive } |
                ForEach-Object { '{0}: active {1} inactive {2}, main {3} {4}' -f $_.Title, $_.Active, $_.Inactive, $script:M.MainTitle, $script:M.MainTitleInactive })
        $bad | Should -BeNullOrEmpty
    }

    It 'keeps the title text of every dialog readable at 4.5:1' {
        $bad = @($script:M.Dialogs | Where-Object { $_.TitleRatio -lt 4.5 } | ForEach-Object { '{0}: {1}' -f $_.Title, $_.TitleRatio })
        $bad | Should -BeNullOrEmpty
    }

    It 'shows bold text on every button of the main window and of every dialog' {
        $all = @($script:M.Main) + @($script:M.Dialogs | ForEach-Object { $_.Buttons })
        $all.Count | Should -BeGreaterThan 5
        $bad = @($all | Where-Object { $_.Weight -lt 600 } | ForEach-Object { '{0} "{1}" weight {2}' -f $_.Window, $_.Text, $_.Weight })
        $bad | Should -BeNullOrEmpty
    }

    It 'gives no button a border thicker than 1 px or in the text color of the theme' {
        $all = @($script:M.Main) + @($script:M.Dialogs | ForEach-Object { $_.Buttons })
        $bad = @($all | Where-Object { $_.BorderThickness -gt 1 -or ($_.BorderThickness -gt 0 -and $_.BorderColor -eq $script:M.ThemeText) } |
                ForEach-Object { '{0} "{1}" border {2}px {3}' -f $_.Window, $_.Text, $_.BorderThickness, $_.BorderColor })
        $bad | Should -BeNullOrEmpty
    }
}
