#Requires -Modules Pester
# Invoke-WpfSta ends a script block that does not finish. These tests create an STA runspace and open no window.

BeforeDiscovery {
    . "$PSScriptRoot\WindowHarness.ps1"
    $script:WindowGuard = Test-WindowTestsAllowed
    if (-not $script:WindowGuard.Allowed) { Write-Host ('Window harness tests skipped: {0}.' -f $script:WindowGuard.Reason) }
}

BeforeAll {
    . "$PSScriptRoot\WindowHarness.ps1"
}

AfterAll { Close-WpfSession }

Describe 'Invoke-WpfSta timeout' -Skip:(-not $script:WindowGuard.Allowed) {
    It 'returns the output of a script block that finishes' {
        Invoke-WpfSta -Script { 2 + 2 } | Should -Be 4
    }

    It 'fails with a clear error when the script block does not finish' {
        { Invoke-WpfSta -TimeoutSeconds 2 -Script { Start-Sleep -Seconds 120 } } | Should -Throw '*did not finish within 2 seconds*'
    }

    It 'closes the stalled session so that the next call starts a new one' {
        $script:HarnessSession | Should -BeNullOrEmpty
        Invoke-WpfSta -Script { 3 + 3 } | Should -Be 6
    }
}
