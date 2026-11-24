#Requires -Modules @{ ModuleName = 'Pester'; ModuleVersion = '5.1.0' }
# Static guard for the grid column headers. It parses the XAML files and creates no WPF control, so it runs on Windows
# PowerShell 5.1 and on PowerShell 7. The window test (HeaderCasing.Window.Tests.ps1) reads the drawn text.
#
# Rules:
#   1. Every column header of every DataGrid is Title Case: each word starts with a capital letter or a digit, and no word
#      is all capitals (a short acronym of the allow-list below is allowed). A header is never typed as a property name
#      (DisplayName) and never typed in capitals.
#   2. The main window has the header style GridHeaderStyle with ContentCharacterCasing Normal, and an implicit
#      DataGridColumnHeader style with the same setter, because MahApps draws headers in capitals by default.
#   3. Every DataGrid gets GridHeaderStyle: through ColumnHeaderStyle on the grid, or through a grid style (also by BasedOn)
#      whose setter for ColumnHeaderStyle names GridHeaderStyle.
#   4. No code adds a grid column or sets a column header.

BeforeDiscovery {
    $script:RootD = Split-Path -Path $PSScriptRoot -Parent
    $script:GridXaml = @(
        (Join-Path $script:RootD 'MainWindow.xaml')
        Get-ChildItem -LiteralPath (Join-Path $script:RootD 'Modules') -Filter '*.xaml' | ForEach-Object FullName
    )
    $script:GridCases = @($script:GridXaml | ForEach-Object { @{ File = $_; Name = (Split-Path -Leaf $_) } })
}

BeforeAll {
    $script:Root = Split-Path -Path $PSScriptRoot -Parent
    $script:Ns = 'http://schemas.microsoft.com/winfx/2006/xaml/presentation'
    $script:XNs = 'http://schemas.microsoft.com/winfx/2006/xaml'
    $script:AllowedAcronym = @('ID', 'URL', 'UNC', 'MSI', 'GUID', 'CPU', 'OS', 'IP')
    $script:GridXaml = @(
        (Join-Path $script:Root 'MainWindow.xaml')
        Get-ChildItem -LiteralPath (Join-Path $script:Root 'Modules') -Filter '*.xaml' | ForEach-Object FullName
    )

    function Read-Xaml {
        param([string]$Path)
        [xml]$doc = Get-Content -LiteralPath $Path -Raw -Encoding UTF8
        $mgr = New-Object System.Xml.XmlNamespaceManager($doc.NameTable)
        $mgr.AddNamespace('w', $script:Ns)
        [pscustomobject]@{ Doc = $doc; Mgr = $mgr }
    }

    # The text of a header must be Title Case words. Returns the reason, or an empty string.
    function Get-HeaderFault {
        param([string]$Header)
        if ([string]::IsNullOrWhiteSpace($Header)) { return 'empty' }
        foreach ($word in ($Header -split '\s+')) {
            if ($word -cnotmatch '^[A-Z0-9]') { return ('word "{0}" does not start with a capital letter' -f $word) }
            if ($word.Length -gt 1 -and $word -cmatch '^[A-Z0-9]+$' -and $word -notmatch '^[0-9]+$' -and $script:AllowedAcronym -cnotcontains $word) { return ('word "{0}" is in capitals' -f $word) }
            if ($word -cmatch '[a-z][A-Z]') { return ('word "{0}" is a joined property name' -f $word) }
        }
        ''
    }

    # The style nodes of MainWindow.xaml by key.
    $script:Main = Read-Xaml -Path (Join-Path $script:Root 'MainWindow.xaml')
    $script:Styles = @{}
    foreach ($node in $script:Main.Doc.SelectNodes("//*[local-name()='Style']")) {
        $key = $node.GetAttribute('Key', $script:XNs)
        if ($key) { $script:Styles[$key] = $node }
    }

    function Get-KeyName {
        param([string]$Reference)
        if ($Reference -match '\{\s*(?:StaticResource|DynamicResource)\s+(?:ResourceKey=)?([A-Za-z0-9_.]+)\s*\}') { return $Matches[1] }
        ''
    }

    # True when the style (or a style it is based on) sets ColumnHeaderStyle to GridHeaderStyle.
    function Test-StyleGivesHeaderStyle {
        param([string]$Key, [int]$Depth = 0)
        if (-not $Key -or $Depth -gt 8 -or -not $script:Styles.ContainsKey($Key)) { return $false }
        $node = $script:Styles[$Key]
        foreach ($setter in @($node.ChildNodes | Where-Object { $_.LocalName -eq 'Setter' -and $_.GetAttribute('Property') -eq 'ColumnHeaderStyle' })) {
            if ((Get-KeyName $setter.GetAttribute('Value')) -eq 'GridHeaderStyle') { return $true }
        }
        Test-StyleGivesHeaderStyle -Key (Get-KeyName $node.GetAttribute('BasedOn')) -Depth ($Depth + 1)
    }
}

Describe 'Grid column header casing' {
    It 'judges sample headers by the rules' {
        Get-HeaderFault 'Display Name' | Should -Be ''
        Get-HeaderFault 'Resource ID' | Should -Be ''
        Get-HeaderFault 'DisplayName' | Should -Not -Be ''
        Get-HeaderFault 'DISPLAYNAME' | Should -Not -Be ''
        Get-HeaderFault 'display name' | Should -Not -Be ''
    }

    It 'has Title Case column headers in <Name>' -ForEach $script:GridCases {
        $x = Read-Xaml -Path $File
        $bad = foreach ($col in $x.Doc.SelectNodes('//w:DataGrid/w:DataGrid.Columns/*', $x.Mgr)) {
            $fault = Get-HeaderFault $col.GetAttribute('Header')
            if ($fault) { 'header "{0}": {1}' -f $col.GetAttribute('Header'), $fault }
        }
        @($bad) | Should -BeNullOrEmpty -Because (@($bad) -join '; ')
    }

    It 'finds both grids and their columns in the XAML files' {
        $grids = 0
        $columns = 0
        foreach ($file in $script:GridXaml) {
            $x = Read-Xaml -Path $file
            $grids += @($x.Doc.SelectNodes('//w:DataGrid', $x.Mgr)).Count
            $columns += @($x.Doc.SelectNodes('//w:DataGrid/w:DataGrid.Columns/*', $x.Mgr)).Count
        }
        $grids | Should -BeGreaterOrEqual 2
        $columns | Should -BeGreaterOrEqual 12
    }

    It 'applies GridHeaderStyle to every grid of <Name>' -ForEach $script:GridCases {
        $x = Read-Xaml -Path $File
        $bad = foreach ($grid in $x.Doc.SelectNodes('//w:DataGrid', $x.Mgr)) {
            $direct = Get-KeyName $grid.GetAttribute('ColumnHeaderStyle')
            $viaStyle = Test-StyleGivesHeaderStyle -Key (Get-KeyName $grid.GetAttribute('Style'))
            if ($direct -ne 'GridHeaderStyle' -and -not $viaStyle) { 'grid "{0}" has no GridHeaderStyle' -f $grid.GetAttribute('Name', $script:XNs) }
        }
        @($bad) | Should -BeNullOrEmpty -Because (@($bad) -join '; ')
    }

    It 'keeps the case of the header text in the main window styles' {
        $header = $script:Styles['GridHeaderStyle']
        $header | Should -Not -BeNullOrEmpty
        @($header.ChildNodes | Where-Object { $_.LocalName -eq 'Setter' -and $_.GetAttribute('Property') -like '*ContentCharacterCasing' -and $_.GetAttribute('Value') -eq 'Normal' }).Count | Should -Be 1
        $implicit = @($script:Main.Doc.SelectNodes("//*[local-name()='Style'][@TargetType='DataGridColumnHeader'][not(@*[local-name()='Key'])]"))
        $implicit.Count | Should -BeGreaterOrEqual 1
        @($implicit[0].ChildNodes | Where-Object { $_.LocalName -eq 'Setter' -and $_.GetAttribute('Property') -like '*ContentCharacterCasing' -and $_.GetAttribute('Value') -eq 'Normal' }).Count | Should -Be 1
    }

    It 'adds no grid column and builds no column header in code' {
        foreach ($file in 'start-detectiontester.ps1', 'Module\DetectionTesterCommon.psm1') {
            $path = Join-Path $script:Root $file
            Test-Path -LiteralPath $path | Should -BeTrue -Because $file
            $text = Get-Content -LiteralPath $path -Raw
            $text | Should -Not -Match 'DataGrid(Text|CheckBox|Template|ComboBox)Column' -Because $file
            $text | Should -Not -Match 'AutoGeneratingColumn|DataGridColumnHeader' -Because $file
        }
    }
}
