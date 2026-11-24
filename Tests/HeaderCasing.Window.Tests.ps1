#Requires -Modules @{ ModuleName = 'Pester'; ModuleVersion = '5.1.0' }
# Window test for the case of the grid column headers (Title Case). The main window is opened off screen in both themes, each module
# XAML (Modules\*.xaml) is loaded into the content host as the shell does it, and the text that each column header draws is read
# from the visual tree. The text on screen must equal the Header text of the column: MahApps draws column headers in capitals
# unless the header style removes it, and the static guard (HeaderCasing.Guard.Tests.ps1) cannot see that.

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
    $script:Design = Get-DesignSize
    Add-Type -AssemblyName PresentationFramework
    $script:ModuleXaml = @(Get-ChildItem -LiteralPath (Join-Path (Split-Path -Parent $PSScriptRoot) 'Modules') -Filter '*.xaml' | ForEach-Object FullName)
    $script:Headers = @{}
    foreach ($dark in $true, $false) {
        $case = @{ Dark = $dark; W = $script:Design.W; H = $script:Design.H; Config = $script:AppConfig; ModuleXaml = $script:ModuleXaml }
        $script:Headers[$(if ($dark) { 'dark' } else { 'light' })] = Invoke-WpfSta -Arguments $case -Script {
            $hcW = $null
            try {
                $hcW = Open-TestWindow -XamlPath (Join-Path $Root $Arguments.Config.Xaml) -Dark $Arguments.Dark -Width $Arguments.W -Height $Arguments.H
                $hcFound = @()
                foreach ($hcFile in $Arguments.ModuleXaml) {
                    [xml]$hcX = Get-Content -LiteralPath $hcFile -Raw -Encoding UTF8
                    $hcW.FindName($Arguments.Config.ContentHost).Content = [System.Windows.Markup.XamlReader]::Load((New-Object System.Xml.XmlNodeReader $hcX))
                    Wait-Layout $hcW
                    foreach ($hcEl in (Get-VisualDescendants $hcW)) {
                        if ($hcEl -isnot [System.Windows.Controls.Primitives.DataGridColumnHeader]) { continue }
                        $hcDrawn = @(Get-VisualDescendants $hcEl | Where-Object { $_ -is [System.Windows.Controls.TextBlock] -and $_.Text } | ForEach-Object { $_.Text })
                        $hcFound += [pscustomobject]@{ View = (Split-Path -Leaf $hcFile); Header = [string]$hcEl.Content; Drawn = ($hcDrawn -join '') }
                    }
                }
                , $hcFound
            }
            finally { if ($hcW) { $hcW.Close() } }
        }
    }
}

AfterAll {
    if ($script:WindowGuard.Allowed) { Close-WpfSession }
}

Describe 'Column headers keep the case of their text <Theme>' -ForEach $script:ThemeCases -Skip:(-not $script:WindowGuard.Allowed) {
    It 'finds the column headers of both grids' {
        $h = @($script:Headers[$Theme])
        foreach ($x in $h) { Write-Host ('{0} {1}: header "{2}" draws "{3}"' -f $Theme, $x.View, $x.Header, $x.Drawn) }
        @($h | Where-Object View -eq 'InstalledApps.xaml').Count | Should -BeGreaterOrEqual 5
        @($h | Where-Object View -eq 'DetectionTester.xaml').Count | Should -BeGreaterOrEqual 7
    }

    It 'draws every column header exactly as its Header text' {
        $bad = @($script:Headers[$Theme] | Where-Object { $_.Drawn -cne $_.Header } | ForEach-Object { '{0}: "{1}" is drawn as "{2}"' -f $_.View, $_.Header, $_.Drawn })
        $bad | Should -BeNullOrEmpty -Because ($bad -join '; ')
    }

    It 'draws the two renamed Installed Applications headers in Title Case' {
        $drawn = @($script:Headers[$Theme] | Where-Object View -eq 'InstalledApps.xaml' | ForEach-Object { $_.Drawn })
        ($drawn -ccontains 'Display Name') | Should -BeTrue -Because ('drawn: ' + ($drawn -join ', '))
        ($drawn -ccontains 'Display Version') | Should -BeTrue -Because ('drawn: ' + ($drawn -join ', '))
    }
}
