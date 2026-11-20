#Requires -Modules @{ ModuleName = 'Pester'; ModuleVersion = '5.1.0' }
# Tooltip guard. It parses the XAML files and the XAML here-strings of the app. It creates no WPF control and opens no window.
#
# Rules:
#   1. Every control in the audited list has a ToolTip.
#   2. The ToolTip text is ASCII, is not longer than 160 characters and ends in a period.
#
# The list holds the controls that Q-103 gave a tooltip. A new input control, check box, radio button, combo box or
# action button belongs in the list when it needs one. Controls that are skipped on purpose: OK and Cancel buttons,
# read-only text boxes, data grid columns and items inside a template.

BeforeDiscovery {
    $script:TtAudited = @(
        'cboType', 'rkvHiveHKLM', 'rkvHiveHKCU', 'rkvView64', 'rkvView32', 'rkvKeyPath', 'rkvValueName', 'rkvOperator', 'rkvExpected',
        'rkHiveHKLM', 'rkHiveHKCU', 'rkView64', 'rkView32', 'rkKeyPath',
        'fileFileName', 'filePropType', 'fileOperator', 'scriptText',
        'cboScope', 'dgApps'
    )
    $script:TtCases = @($script:TtAudited | ForEach-Object { @{ Name = $_ } })
}

BeforeAll {
    $script:TtRepoRoot = Split-Path -Parent $PSScriptRoot
    $script:TtNamespace = 'http://schemas.microsoft.com/winfx/2006/xaml/presentation'
    $script:TtKeyNamespace = 'http://schemas.microsoft.com/winfx/2006/xaml'
    $script:TtExcluded = '^(\.claude|\.git|Tests|tools|Logs|Lib)([\\/]|$)'

    function Get-TtRelativePath {
        param([string]$FullName)
        return $FullName.Substring($script:TtRepoRoot.Length).TrimStart('\', '/')
    }

    # The XAML here-strings of a script.
    function Get-TtHereString {
        param([string]$Text)
        $tokens = $null; $errors = $null
        $ast = [System.Management.Automation.Language.Parser]::ParseInput($Text, [ref]$tokens, [ref]$errors)
        $nodes = $ast.FindAll({ param($n) $n -is [System.Management.Automation.Language.StringConstantExpressionAst] -and $n.StringConstantType -in 'SingleQuotedHereString', 'DoubleQuotedHereString' }, $true)
        foreach ($n in @($nodes)) {
            if ($n.Value -notmatch ('^\s*<[^>]*xmlns="' + [regex]::Escape($script:TtNamespace) + '"')) { continue }
            $n.Value
        }
    }

    # Every XAML document of the app: the XAML files and the XAML here-strings of the scripts.
    function Get-TtDocument {
        $docs = @()
        $files = @(Get-ChildItem -LiteralPath $script:TtRepoRoot -File -Recurse | Where-Object { (Get-TtRelativePath $_.FullName) -notmatch $script:TtExcluded })
        foreach ($f in @($files | Where-Object { $_.Extension -eq '.xaml' })) {
            $docs += @{ Source = (Get-TtRelativePath $f.FullName); Text = (Get-Content -LiteralPath $f.FullName -Raw) }
        }
        foreach ($f in @($files | Where-Object { $_.Extension -in '.ps1', '.psm1' -and $_.Name -notlike '*.Tests.ps1' })) {
            $i = 0
            foreach ($h in @(Get-TtHereString -Text (Get-Content -LiteralPath $f.FullName -Raw))) {
                $i++
                $docs += @{ Source = ('{0} here-string {1}' -f (Get-TtRelativePath $f.FullName), $i); Text = $h }
            }
        }
        return $docs
    }

    # The elements of every document, by x:Name.
    $script:TtElement = @{}
    foreach ($d in @(Get-TtDocument)) {
        [xml]$x = $d.Text
        foreach ($el in @($x.SelectNodes('//*'))) {
            $n = $el.GetAttribute('Name', $script:TtKeyNamespace)
            if (-not $n) { $n = $el.GetAttribute('Name') }
            if ($n) { $script:TtElement[$n] = $el }
        }
    }
}

Describe 'Tooltips of the audited controls' {
    It 'finds the control <Name>' -ForEach $script:TtCases {
        $script:TtElement.ContainsKey($Name) | Should -BeTrue -Because "$Name is in the audited list"
    }

    It 'gives <Name> a tooltip' -ForEach $script:TtCases {
        $tip = $script:TtElement[$Name].GetAttribute('ToolTip')
        $tip | Should -Not -BeNullOrEmpty -Because "$Name needs a ToolTip"
    }

    It 'writes the tooltip of <Name> in ASCII, in one sentence that ends in a period' -ForEach $script:TtCases {
        $tip = $script:TtElement[$Name].GetAttribute('ToolTip')
        $tip | Should -Match '^[\x20-\x7e]+$' -Because "$Name tooltip is ASCII"
        $tip | Should -Match '\.$' -Because "$Name tooltip ends in a period"
        $tip | Should -Not -Match '\.\s+\S' -Because "$Name tooltip is one sentence"
        $tip.Length | Should -BeLessOrEqual 160 -Because "$Name tooltip is short"
    }
}
