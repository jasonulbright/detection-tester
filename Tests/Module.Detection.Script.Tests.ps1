#Requires -Modules @{ ModuleName = 'Pester'; ModuleVersion = '5.1.0' }
# Script detection.

BeforeAll {
    . "$PSScriptRoot\TestHelpers.ps1"
    Import-AppModule
    Mock Write-Log {} -ModuleName DetectionTesterCommon
}
Describe "Test-ScriptDetection" {
    It "detects when script produces output" {
        $result = Test-ScriptDetection -ScriptText 'Write-Output "found"'

        $result.Detected | Should -Be $true
        $result.ScriptOutput | Should -Match "found"
    }

    It "returns not detected when script produces no output" {
        $result = Test-ScriptDetection -ScriptText '$x = 1'

        $result.Detected | Should -Be $false
    }

    It "returns not detected on script error with no stdout" {
        $result = Test-ScriptDetection -ScriptText 'Write-Error "fail"'

        $result.Detected | Should -Be $false
    }
}

Describe 'Codex review fixes' {
    Context "Test-ScriptDetection: isolation + timeout" {
        It "does not terminate the host process when script calls exit" {
            # If isolation is broken, this `It` block never returns (the test
            # process exits). Pester reaching the assertion means we survived.
            $r = Test-ScriptDetection -ScriptText 'exit 7'
            $r.Detected | Should -Be $false
            $r.TimedOut | Should -Be $false
        }
        It "returns TimedOut=true on a sleeping script (with short timeout)" {
            $r = Test-ScriptDetection -ScriptText 'Start-Sleep -Seconds 30; "late"' -TimeoutSeconds 2
            $r.TimedOut | Should -Be $true
            $r.Detected | Should -Be $false
        }
        It "still detects normal output" {
            $r = Test-ScriptDetection -ScriptText 'Write-Output "hello"' -TimeoutSeconds 5
            $r.Detected | Should -Be $true
            $r.TimedOut | Should -Be $false
            $r.ScriptOutput | Should -Match 'hello'
        }
    }
}

Describe 'Test-ScriptDetection output rules' {
    It 'gives Detected=<Detected> for the script <Script>' -ForEach @(
        @{ Script = 'Write-Output "found"';         Detected = $true }
        @{ Script = '"x"';                          Detected = $true }
        @{ Script = '0';                            Detected = $true }
        @{ Script = '$false';                       Detected = $true }
        @{ Script = '1,2,3';                        Detected = $true }
        @{ Script = '@{a=1}';                       Detected = $true }
        @{ Script = '$x = 1';                       Detected = $false }
        @{ Script = '$null';                        Detected = $false }
        @{ Script = '"" ';                          Detected = $false }
        @{ Script = '"   "';                        Detected = $false }
        @{ Script = 'Write-Host "host text only"';  Detected = $false }
        @{ Script = 'Write-Warning "w"';            Detected = $false }
        @{ Script = 'Write-Error "e"';              Detected = $false }
        @{ Script = '   ';                          Detected = $false }
        @{ Script = 'Get-Date | Out-Null';          Detected = $false }
        @{ Script = 'throw "boom"';                 Detected = $false }
        @{ Script = '"x"; throw "late"';            Detected = $false }
        @{ Script = 'exit 7';                       Detected = $false }
        @{ Script = '1 +';                          Detected = $false }
    ) {
        $r = Test-ScriptDetection -ScriptText $Script -TimeoutSeconds 20
        $r.Detected | Should -Be $Detected
        $r.TimedOut | Should -BeFalse
        $r.Type | Should -Be 'Script'
        $r.Target | Should -Be '(script)'
        $r.Operator | Should -Be 'Script'
        $r.ExpectedValue | Should -Be '(non-empty output)'
    }

    It 'returns the output lines joined by a line break' {
        $r = Test-ScriptDetection -ScriptText '"one"; "two"; "three"'
        $r.ScriptOutput | Should -Be "one`r`ntwo`r`nthree"
        $r.ActualValue | Should -Be $r.ScriptOutput
    }

    It 'keeps non-ASCII output intact' {
        $r = Test-ScriptDetection -ScriptText '[string][char]0xE9 + [char]0x20AC'
        $r.ScriptOutput | Should -Be "$([char]0xE9)$([char]0x20AC)"
    }

    It 'returns the whole output and only a 200 character preview in the details' {
        $r = Test-ScriptDetection -ScriptText '"a" * 5000'
        $r.ScriptOutput.Length | Should -Be 5000
        $r.Details | Should -Be ('Script produced output (DETECTED): ' + ('a' * 200))
    }

    It 'shows the error text next to "no output" when the script only writes an error' {
        $r = Test-ScriptDetection -ScriptText 'Write-Error "disk missing"'
        $r.Details | Should -Be 'Script produced no output (NOT DETECTED) | Errors: disk missing'
    }

    It 'shows a fixed message and no exception text when the script throws' {
        $r = Test-ScriptDetection -ScriptText 'throw "boom"'
        $r.Details | Should -Be 'The script could not run. See the log for details.'
        $r.Details | Should -Not -Match 'boom|Exception'
        $r.ScriptOutput | Should -BeNullOrEmpty
    }

    It 'rejects a script language other than PowerShell: <Language>' -ForEach @(
        @{ Language = 'VBScript' }, @{ Language = 'JScript' }, @{ Language = 'cmd' }, @{ Language = '' }
    ) {
        $r = Test-ScriptDetection -ScriptText '"x"' -ScriptLanguage $Language
        $r.Detected | Should -BeFalse
        $r.Details | Should -Be "Unsupported script language: $Language (only PowerShell supported)"
    }

    It 'accepts the language name in any case' {
        (Test-ScriptDetection -ScriptText '"x"' -ScriptLanguage 'powershell').Detected | Should -BeTrue
    }

    It 'runs each script in its own session and leaves the host session unchanged' {
        $before = (Get-Location).Path
        $null = Test-ScriptDetection -ScriptText '$global:ZZ_DT_LEAK = "leak"; $script:ZZ_DT_LEAK2 = 1; Set-Location $env:windir; "x"'
        Get-Variable -Name ZZ_DT_LEAK -Scope Global -ErrorAction SilentlyContinue | Should -BeNullOrEmpty
        (Get-Location).Path | Should -Be $before
        $second = Test-ScriptDetection -ScriptText 'if (Test-Path variable:global:ZZ_DT_LEAK) { "seen" }'
        $second.Detected | Should -BeFalse
    }
}

Describe 'Test-ScriptDetection timeout' {
    It 'stops a script that sleeps past the timeout and does not return its earlier output' {
        $sw = [Diagnostics.Stopwatch]::StartNew()
        $r = Test-ScriptDetection -ScriptText '"early"; Start-Sleep -Seconds 30; "late"' -TimeoutSeconds 1
        $sw.Stop()
        $r.TimedOut | Should -BeTrue
        $r.Detected | Should -BeFalse
        $r.ScriptOutput | Should -BeNullOrEmpty
        $r.Details | Should -Be 'Script timed out after 1s (no output captured); treated as NOT DETECTED.'
        $sw.Elapsed.TotalSeconds | Should -BeLessThan 20
    }

    It 'stops a script that never ends' {
        $r = Test-ScriptDetection -ScriptText 'while ($true) { }' -TimeoutSeconds 1
        $r.TimedOut | Should -BeTrue
        $r.Detected | Should -BeFalse
    }

    It 'waits at least one second and reports the real limit for a timeout of <Timeout>' -ForEach @(
        @{ Timeout = 0 }, @{ Timeout = -5 }
    ) {
        $r = Test-ScriptDetection -ScriptText 'Start-Sleep -Seconds 30' -TimeoutSeconds $Timeout
        $r.TimedOut | Should -BeTrue
        $r.Details | Should -BeLike 'Script timed out after 1s *'
    }

    It 'returns the output of a quick script for a timeout of <Timeout>' -ForEach @(@{ Timeout = 0 }, @{ Timeout = -5 }) {
        $r = Test-ScriptDetection -ScriptText '"ok"' -TimeoutSeconds $Timeout
        $r.TimedOut | Should -BeFalse
        $r.Detected | Should -BeTrue
        $r.ScriptOutput | Should -Be 'ok'
    }

    It 'returns normal output well before a long timeout' {
        $sw = [Diagnostics.Stopwatch]::StartNew()
        $r = Test-ScriptDetection -ScriptText '"quick"' -TimeoutSeconds 60
        $sw.Stop()
        $r.Detected | Should -BeTrue
        $sw.Elapsed.TotalSeconds | Should -BeLessThan 30
    }
}
