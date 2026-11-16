#Requires -Modules @{ ModuleName = 'Pester'; ModuleVersion = '5.1.0' }
# Headless test for the compact 1080p density of the main window (XAML only, no window is opened).
# The shared style values are asserted in one place. Density.Window.Tests.ps1 measures them on screen.

BeforeAll {
    $script:Root = Split-Path -Path $PSScriptRoot -Parent
    $script:Xaml = [xml](Get-Content -LiteralPath (Join-Path $script:Root 'MainWindow.xaml') -Raw)
    $script:Ns = New-Object System.Xml.XmlNamespaceManager($script:Xaml.NameTable)
    $script:Ns.AddNamespace('w', 'http://schemas.microsoft.com/winfx/2006/xaml/presentation')
    $script:Ns.AddNamespace('x', 'http://schemas.microsoft.com/winfx/2006/xaml')
    $script:Modules = @(foreach ($f in 'DetectionTester.xaml', 'InstalledApps.xaml') { [xml](Get-Content -LiteralPath (Join-Path $script:Root "Modules\$f") -Raw) })
    function Get-ModuleNode {
        param([string]$XPath)
        foreach ($m in $script:Modules) {
            $ns = New-Object System.Xml.XmlNamespaceManager($m.NameTable)
            $ns.AddNamespace('w', 'http://schemas.microsoft.com/winfx/2006/xaml/presentation')
            $ns.AddNamespace('x', 'http://schemas.microsoft.com/winfx/2006/xaml')
            $m.SelectNodes($XPath, $ns)
        }
    }    function Get-StyleValue {
        param([string]$Key, [string]$Property)
        $style = $script:Xaml.SelectSingleNode("//w:Style[@x:Key='$Key']", $script:Ns)
        $style | Should -Not -BeNullOrEmpty -Because $Key
        $setter = $style.SelectSingleNode("w:Setter[@Property='$Property']", $script:Ns)
        if ($setter) { $setter.GetAttribute('Value') } else { $null }
    }
}

Describe 'Compact density at 1080p' {
    It 'opens at the design size of a maximized 1080p work area and keeps a minimum that fits 1024x720' {
        $w = $script:Xaml.DocumentElement
        [int]$w.GetAttribute('Width') | Should -Be 1920
        [int]$w.GetAttribute('Height') | Should -Be 1032
        [int]$w.GetAttribute('MinWidth') | Should -BeLessOrEqual 1024
        [int]$w.GetAttribute('MinHeight') | Should -BeLessOrEqual 720
    }

    It 'sizes the sidebar view buttons between 34 and 36 px' {
        [double](Get-StyleValue 'ViewButtonStyle' 'Height') | Should -BeIn 34, 35, 36
    }

    It 'sizes the inline buttons, the primary button and the implicit combo box between 26 and 28 px' {
        foreach ($key in 'SquareInlineButtonStyle', 'PrimaryButtonStyle') {
            [double](Get-StyleValue $key 'Height') | Should -BeGreaterOrEqual 26 -Because $key
            [double](Get-StyleValue $key 'Height') | Should -BeLessOrEqual 28 -Because $key
        }
        $combo = $script:Xaml.SelectSingleNode("//w:Style[@TargetType='ComboBox' and not(@x:Key)]", $script:Ns)
        [double]$combo.SelectSingleNode("w:Setter[@Property='Height']", $script:Ns).GetAttribute('Value') | Should -BeIn 26, 27, 28
    }

    It 'sizes every single-line text box between 26 and 28 px' {
        $boxes = @(Get-ModuleNode '//w:TextBox[@Height and not(@x:Name="scriptText")]')
        $boxes.Count | Should -BeGreaterThan 5
        foreach ($tb in $boxes) {
            [double]$tb.GetAttribute('Height') | Should -BeIn 26, 27, 28 -Because $tb.GetAttribute('Name', 'http://schemas.microsoft.com/winfx/2006/xaml')
        }
    }

    It 'gives every grid rows of 24 to 26 px and a 26 px column header from the one shared style' {
        [double](Get-StyleValue 'HistoryGridStyle' 'RowHeight') | Should -BeGreaterOrEqual 24
        [double](Get-StyleValue 'HistoryGridStyle' 'RowHeight') | Should -BeLessOrEqual 26
        [double](Get-StyleValue 'HistoryGridStyle' 'ColumnHeaderHeight') | Should -Be 26
        $grids = @(Get-ModuleNode '//w:DataGrid')
        $grids.Count | Should -BeGreaterThan 1
        foreach ($g in $grids) {
            $g.GetAttribute('RowHeight') | Should -BeNullOrEmpty -Because 'the row height has one place: HistoryGridStyle'
            $g.GetAttribute('Style') | Should -Match 'HistoryGridStyle' -Because $g.GetAttribute('Name', 'http://schemas.microsoft.com/winfx/2006/xaml')
        }
    }
}
