#Requires -Modules Pester
# The window guard. These tests open no window and create no runspace.

BeforeAll {
    . "$PSScriptRoot\WindowHarness.ps1"
    $script:Other = 'OTHERPC'
}

Describe 'Test-WindowTestsAllowed' {
    It 'refuses the dev PC "<Name>" even with the opt-in' -ForEach @(
        @{ Name = 'AYW' }, @{ Name = 'ayw' }, @{ Name = 'Ayw' }
    ) {
        $g = Test-WindowTestsAllowed -ComputerName $Name -OptIn '1'
        $g.Allowed | Should -BeFalse
        $g.Reason | Should -BeLike '*AYW*'
    }

    It 'allows another machine with the opt-in' {
        $g = Test-WindowTestsAllowed -ComputerName $script:Other -OptIn '1'
        $g.Allowed | Should -BeTrue
        $g.Reason | Should -Be ''
    }

    It 'refuses another machine when the opt-in is "<OptIn>"' -ForEach @(
        @{ OptIn = '' }, @{ OptIn = '0' }, @{ OptIn = 'true' }, @{ OptIn = $null }
    ) {
        $g = Test-WindowTestsAllowed -ComputerName $script:Other -OptIn $OptIn
        $g.Allowed | Should -BeFalse
        $g.Reason | Should -Be 'SUITE_ALLOW_WINDOW_TESTS is not 1'
    }

    It 'gives a reason whenever it refuses' -ForEach @(
        @{ Name = 'AYW'; OptIn = '1' }, @{ Name = 'CM01'; OptIn = '' }
    ) {
        (Test-WindowTestsAllowed -ComputerName $Name -OptIn $OptIn).Reason | Should -Not -BeNullOrEmpty
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
