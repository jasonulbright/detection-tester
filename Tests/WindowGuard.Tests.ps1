#Requires -Modules Pester
# The window guard. These tests open no window and create no runspace.

BeforeAll {
    . "$PSScriptRoot\WindowHarness.ps1"
}

Describe 'Test-WindowTestsAllowed' {
    It 'allows the run with the opt-in' {
        $g = Test-WindowTestsAllowed -OptIn '1'
        $g.Allowed | Should -BeTrue
        $g.Reason | Should -Be ''
    }

    It 'refuses the run when the opt-in is "<OptIn>"' -ForEach @(
        @{ OptIn = '' }, @{ OptIn = '0' }, @{ OptIn = 'true' }, @{ OptIn = '1 ' }, @{ OptIn = $null }
    ) {
        $g = Test-WindowTestsAllowed -OptIn $OptIn
        $g.Allowed | Should -BeFalse
        $g.Reason | Should -Be 'SUITE_ALLOW_WINDOW_TESTS is not 1'
    }
}

Describe 'Test-WindowTestsAllowed environment default' {
    BeforeEach {
        $script:savedOptIn = $env:SUITE_ALLOW_WINDOW_TESTS
    }

    AfterEach {
        if ($null -eq $script:savedOptIn) { Remove-Item -LiteralPath 'Env:\SUITE_ALLOW_WINDOW_TESTS' -ErrorAction SilentlyContinue }
        else { $env:SUITE_ALLOW_WINDOW_TESTS = $script:savedOptIn }
    }

    It 'allows the run when the environment value is 1' {
        $env:SUITE_ALLOW_WINDOW_TESTS = '1'
        $g = Test-WindowTestsAllowed
        $g.Allowed | Should -BeTrue
        $g.Reason | Should -Be ''
    }

    It 'refuses the run when the environment value is "<Value>"' -ForEach @(
        @{ Value = '0' }, @{ Value = 'true' }
    ) {
        $env:SUITE_ALLOW_WINDOW_TESTS = $Value
        $g = Test-WindowTestsAllowed
        $g.Allowed | Should -BeFalse
        $g.Reason | Should -Be 'SUITE_ALLOW_WINDOW_TESTS is not 1'
    }

    It 'refuses the run when the environment value is not set' {
        Remove-Item -LiteralPath 'Env:\SUITE_ALLOW_WINDOW_TESTS' -ErrorAction SilentlyContinue
        $g = Test-WindowTestsAllowed
        $g.Allowed | Should -BeFalse
        $g.Reason | Should -Be 'SUITE_ALLOW_WINDOW_TESTS is not 1'
    }
}

Describe 'Test-WindowTestsAllowed signature' {
    It 'has the OptIn parameter and no ComputerName parameter' {
        $names = @((Get-Command Test-WindowTestsAllowed).Parameters.Keys | Where-Object { $_ -notin [System.Management.Automation.Cmdlet]::CommonParameters })
        $names | Should -Be @('OptIn')
    }

    It 'decides on the opt-in only' {
        (Test-WindowTestsAllowed -OptIn '1').Allowed | Should -BeTrue
        (Test-WindowTestsAllowed -OptIn '').Allowed | Should -BeFalse
    }
}

Describe 'Window test files' {
    It 'checks the guard in every test file that opens a window' {
        $self = 'WindowGuard.Tests.ps1'
        $pattern = 'Invoke-WpfSta|New-WpfWindow|Open-TestWindow|\.ShowDialog\(|\.Show\(\)'
        $files = @(Get-ChildItem -LiteralPath $PSScriptRoot -Filter '*.Tests.ps1' -File | Where-Object { $_.Name -ne $self })
        $files.Count | Should -BeGreaterThan 0
        $missing = @($files | Where-Object {
                $text = Get-Content -LiteralPath $_.FullName -Raw
                $text -match $pattern -and $text -notmatch 'Test-WindowTestsAllowed'
            } | ForEach-Object Name)
        $missing | Should -BeNullOrEmpty
    }
}
