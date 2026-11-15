#Requires -Modules @{ ModuleName = 'Pester'; ModuleVersion = '5.1.0' }
# Headless guard for the cleanup of stray windows. It opens no window and creates no runspace.
#   - Invoke-WpfSta of the harness runs Close-SuiteOpenWindow in its finally block, so a test that fails before its own Close leaves no window.
#   - A window test file opens a window (Open-TestWindow, .Show(), .ShowDialog()) only inside a try statement that has a finally block.
#     The one exception is the persistent window of the handler tests ($global:HW), which an AfterAll closes.

BeforeDiscovery {
    $script:TestsDir = $PSScriptRoot
    $script:WindowTestFile = @(Get-ChildItem -LiteralPath $PSScriptRoot -Filter '*.Window.Tests.ps1' -File) + @(Get-Item -LiteralPath (Join-Path $PSScriptRoot 'Window.Tests.ps1'))
}

BeforeAll {
    function Get-Ast {
        param([string]$Path)
        $e = $null
        [System.Management.Automation.Language.Parser]::ParseFile($Path, [ref]$null, [ref]$e)
    }
}

Describe 'Window cleanup guard' {
    It 'closes the open windows in the finally block of Invoke-WpfSta' {
        $ast = Get-Ast (Join-Path $PSScriptRoot 'WindowHarness.ps1')
        $fn = $ast.Find({ param($n) $n -is [System.Management.Automation.Language.FunctionDefinitionAst] -and $n.Name -eq 'Invoke-WpfSta' }, $true)
        $fn | Should -Not -BeNullOrEmpty
        $try = $fn.Body.FindAll({ param($n) $n -is [System.Management.Automation.Language.TryStatementAst] -and $n.Finally }, $true)
        @($try | Where-Object { $_.Finally.Extent.Text -match 'Close-SuiteOpenWindow' }).Count | Should -BeGreaterThan 0
    }

    It 'opens a window in <Name> only inside a try statement with a finally block' -ForEach @($script:WindowTestFile | ForEach-Object { @{ Name = $_.Name; Path = $_.FullName } }) {
        $ast = Get-Ast $Path
        $calls = @($ast.FindAll({
                    param($n)
                    ($n -is [System.Management.Automation.Language.CommandAst] -and $n.GetCommandName() -eq 'Open-TestWindow') -or
                    ($n -is [System.Management.Automation.Language.InvokeMemberExpressionAst] -and $n.Member.Value -in 'Show', 'ShowDialog')
                }, $true))
        $bad = @()
        foreach ($c in $calls) {
            $p = $c.Parent
            $guarded = $false
            while ($p) {
                if ($p -is [System.Management.Automation.Language.AssignmentStatementAst] -and $p.Left.Extent.Text -eq '$global:HW') { $guarded = $true; break }
                if ($p -is [System.Management.Automation.Language.TryStatementAst] -and $p.Finally) { $guarded = $true; break }
                $p = $p.Parent
            }
            if (-not $guarded) {
                # The window opens in the statement just before a try statement that has a finally block.
                $stmt = $c
                while ($stmt.Parent -and $stmt.Parent -isnot [System.Management.Automation.Language.NamedBlockAst] -and $stmt.Parent -isnot [System.Management.Automation.Language.StatementBlockAst]) { $stmt = $stmt.Parent }
                if ($stmt.Parent) {
                    $sib = @($stmt.Parent.Statements)
                    $at = [array]::IndexOf($sib, $stmt)
                    if ($at -ge 0 -and $at + 1 -lt $sib.Count -and $sib[$at + 1] -is [System.Management.Automation.Language.TryStatementAst] -and $sib[$at + 1].Finally) { $guarded = $true }
                }
            }            if (-not $guarded) { $bad += ('line {0}: {1}' -f $c.Extent.StartLineNumber, $c.Extent.Text.Split("`n")[0].Trim()) }
        }
        $bad | Should -BeNullOrEmpty -Because 'a window that stays open on a failure stays on the desktop of the lab computer'
    }

    It 'closes the persistent window of the handler tests in an AfterAll' {
        $text = Get-Content -LiteralPath (Join-Path $PSScriptRoot 'Window.Tests.ps1') -Raw
        $text | Should -Match '(?s)AfterAll\s*\{[^}]*\$global:HW\.Close\(\)'
    }
}