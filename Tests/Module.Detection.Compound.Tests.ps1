#Requires -Modules @{ ModuleName = 'Pester'; ModuleVersion = '5.1.0' }
# Compound detection.

BeforeAll {
    . "$PSScriptRoot\TestHelpers.ps1"
    Import-AppModule
    Mock Write-Log {} -ModuleName DetectionTesterCommon
}
Describe "Test-CompoundDetection" {
    It "And connector: all must pass" {
        $clauses = @(
            @{ Type = 'RegistryKey'; RegistryKeyRelative = 'SOFTWARE\Microsoft\Windows NT\CurrentVersion'; Is64Bit = $true }
            @{ Type = 'File'; FilePath = 'C:\Windows'; FileName = 'explorer.exe'; PropertyType = 'Existence' }
        )
        $result = Test-CompoundDetection -Connector 'And' -Clauses $clauses

        $result.Detected | Should -Be $true
        $result.ClauseResults.Count | Should -Be 2
    }

    It "And connector: fails if one clause fails" {
        $clauses = @(
            @{ Type = 'RegistryKey'; RegistryKeyRelative = 'SOFTWARE\Microsoft\Windows NT\CurrentVersion'; Is64Bit = $true }
            @{ Type = 'RegistryKey'; RegistryKeyRelative = 'SOFTWARE\NonExistentKey12345'; Is64Bit = $true }
        )
        $result = Test-CompoundDetection -Connector 'And' -Clauses $clauses

        $result.Detected | Should -Be $false
    }

    It "Or connector: passes if one clause passes" {
        $clauses = @(
            @{ Type = 'RegistryKey'; RegistryKeyRelative = 'SOFTWARE\NonExistentKey12345'; Is64Bit = $true }
            @{ Type = 'RegistryKey'; RegistryKeyRelative = 'SOFTWARE\Microsoft\Windows NT\CurrentVersion'; Is64Bit = $true }
        )
        $result = Test-CompoundDetection -Connector 'Or' -Clauses $clauses

        $result.Detected | Should -Be $true
    }
}

Describe 'Test-CompoundDetection results' {
    BeforeAll {
        $script:Key = New-TestRegistryKey -Value @{ Ver = '9.9.9'; DisplayVersion = '2.0.0'; Zero = @{ Data = 0; Kind = 'DWord' }; Flag = @{ Data = 1; Kind = 'DWord' } }
        $script:Dir = Join-Path $TestDrive 'cmp'
        New-Item -ItemType Directory -Path $script:Dir -Force | Out-Null
        Set-Content -LiteralPath (Join-Path $script:Dir 'present.txt') -Value 'x'
        $script:Yes = @{ Type = 'RegistryKey'; RegistryKeyRelative = $script:Key; Hive = 'CurrentUser' }
        $script:No = @{ Type = 'RegistryKey'; RegistryKeyRelative = ($script:Key + '\missing'); Hive = 'CurrentUser' }
    }
    AfterAll { Remove-TestRegistryKey -Relative $script:Key }

    It 'gives <Connector> = <Detected> for clauses <Pattern>' -ForEach @(
        @{ Connector = 'And'; Pattern = 'yes,yes';    Detected = $true;  Passed = 2 }
        @{ Connector = 'And'; Pattern = 'yes,no';     Detected = $false; Passed = 1 }
        @{ Connector = 'And'; Pattern = 'no,yes';     Detected = $false; Passed = 1 }
        @{ Connector = 'And'; Pattern = 'no,no';      Detected = $false; Passed = 0 }
        @{ Connector = 'And'; Pattern = 'yes';        Detected = $true;  Passed = 1 }
        @{ Connector = 'And'; Pattern = 'no';         Detected = $false; Passed = 0 }
        @{ Connector = 'Or';  Pattern = 'yes,yes';    Detected = $true;  Passed = 2 }
        @{ Connector = 'Or';  Pattern = 'no,yes';     Detected = $true;  Passed = 1 }
        @{ Connector = 'Or';  Pattern = 'yes,no,no';  Detected = $true;  Passed = 1 }
        @{ Connector = 'Or';  Pattern = 'no,no';      Detected = $false; Passed = 0 }
        @{ Connector = 'AND'; Pattern = 'yes,yes';    Detected = $true;  Passed = 2 }
        @{ Connector = 'or';  Pattern = 'no,yes';     Detected = $true;  Passed = 1 }
    ) {
        $clauses = @($Pattern -split ',' | ForEach-Object { if ($_ -eq 'yes') { $script:Yes } else { $script:No } })
        $r = Test-CompoundDetection -Connector $Connector -Clauses $clauses
        $r.Detected | Should -Be $Detected
        $r.ClauseResults.Count | Should -Be $clauses.Count
        $r.ActualValue | Should -Be ('{0}/{1} passed' -f $Passed, $clauses.Count)
        $r.Details | Should -Be ('{0}: {1}/{2} clauses passed - {3}' -f $Connector, $Passed, $clauses.Count, $(if ($Detected) { 'DETECTED' } else { 'NOT DETECTED' }))
        $r.Type | Should -Be 'Compound'
        $r.Operator | Should -Be $Connector
        $r.Target | Should -Be ('({0}: {1} clauses)' -f $Connector, $clauses.Count)
    }

    It 'lists the results of the clauses in the order of the clauses' {
        $r = Test-CompoundDetection -Connector And -Clauses @($script:No, $script:Yes)
        $r.ClauseResults[0].Detected | Should -BeFalse
        $r.ClauseResults[1].Detected | Should -BeTrue
    }

    It 'does not detect with the connector "<Connector>" and says that the connector is not supported' -ForEach @(
        @{ Connector = 'Xor' }, @{ Connector = 'All' }, @{ Connector = 'Any' }
    ) {
        $r = Test-CompoundDetection -Connector $Connector -Clauses @($script:Yes)
        $r.Detected | Should -BeFalse
        $r.Details | Should -BeLike "Unsupported connector*"
    }

    It 'evaluates a RegistryKeyValue clause and defaults the type, the value name, the operator and the hive' {
        # No ValueName: the default is DisplayVersion. Ver holds a different value on purpose.
        $clause = @{ RegistryKeyRelative = $script:Key; Hive = 'CurrentUser'; ExpectedValue = '2.0.0' }
        $r = Test-CompoundDetection -Connector And -Clauses @($clause)
        $r.ClauseResults[0].Type | Should -Be 'RegistryKeyValue'
        $r.ClauseResults[0].Detected | Should -BeTrue
        $r.ClauseResults[0].Operator | Should -Be 'IsEquals'
    }

    It 'takes the expected value from DisplayVersion when ExpectedValue is missing' {
        $clause = @{ Type = 'RegistryKeyValue'; RegistryKeyRelative = $script:Key; Hive = 'CurrentUser'; ValueName = 'DisplayVersion'; DisplayVersion = '2.0.0' }
        (Test-CompoundDetection -Connector And -Clauses @($clause)).Detected | Should -BeTrue
    }

    It 'evaluates an expected value of 0 as the number 0' {
        $clause = @{ Type = 'RegistryKeyValue'; RegistryKeyRelative = $script:Key; Hive = 'CurrentUser'; ValueName = 'Zero'; ExpectedValue = 0 }
        $r = Test-CompoundDetection -Connector And -Clauses @($clause)
        $r.ClauseResults[0].ExpectedValue | Should -Be '0'
        $r.ClauseResults[0].Detected | Should -BeTrue
    }

    It 'does not treat an expected value of 0 as a match for a value of 1' {
        $clause = @{ Type = 'RegistryKeyValue'; RegistryKeyRelative = $script:Key; Hive = 'CurrentUser'; ValueName = 'Flag'; ExpectedValue = 0 }
        (Test-CompoundDetection -Connector And -Clauses @($clause)).Detected | Should -BeFalse
    }

    It 'passes the operator and the property type of a RegistryKeyValue clause' {
        $clause = @{ Type = 'RegistryKeyValue'; RegistryKeyRelative = $script:Key; Hive = 'CurrentUser'; ValueName = 'Flag'; ExpectedValue = '1'; Operator = 'GreaterThan'; PropertyType = 'DWORD' }
        (Test-CompoundDetection -Connector And -Clauses @($clause)).Detected | Should -BeFalse
        $clause.Operator = 'GreaterEquals'
        (Test-CompoundDetection -Connector And -Clauses @($clause)).Detected | Should -BeTrue
    }

    It 'reads a clause that is an object instead of a hashtable' {
        $clause = [pscustomobject]@{ Type = 'RegistryKey'; RegistryKeyRelative = $script:Key; Hive = 'CurrentUser' }
        (Test-CompoundDetection -Connector And -Clauses @($clause)).Detected | Should -BeTrue
    }

    It 'evaluates a File clause with the defaults Existence and GreaterEquals' {
        $r = Test-CompoundDetection -Connector And -Clauses @(@{ Type = 'File'; FilePath = $script:Dir; FileName = 'present.txt' })
        $r.Detected | Should -BeTrue
        $r.ClauseResults[0].Operator | Should -Be 'Existence'
        (Test-CompoundDetection -Connector And -Clauses @(@{ Type = 'File'; FilePath = $script:Dir; FileName = 'absent.txt' })).Detected | Should -BeFalse
    }

    It 'evaluates a Script clause' {
        (Test-CompoundDetection -Connector Or -Clauses @(@{ Type = 'Script'; ScriptText = '"found"' })).Detected | Should -BeTrue
        (Test-CompoundDetection -Connector Or -Clauses @(@{ Type = 'Script'; ScriptText = '$null' })).Detected | Should -BeFalse
    }

    It 'treats a clause of an unknown type as not detected and names the type' {
        $r = Test-CompoundDetection -Connector Or -Clauses @(@{ Type = 'WmiQuery' }, $script:Yes)
        $r.ClauseResults[0].Detected | Should -BeFalse
        $r.ClauseResults[0].Details | Should -Be 'Unknown clause type: WmiQuery'
        $r.Detected | Should -BeTrue
    }

    It 'does not accept a compound clause inside a compound clause' {
        $r = Test-CompoundDetection -Connector And -Clauses @(@{ Type = 'Compound'; Connector = 'And'; Clauses = @($script:Yes) })
        $r.Detected | Should -BeFalse
        $r.ClauseResults[0].Details | Should -Be 'Unknown clause type: Compound'
    }

    # A clause that cannot run (for example a manifest that left the key path empty) must fail that clause
    # and keep the result of the other clauses. The shell calls this function from a click handler.
    It 'fails only the clause that cannot run: <Case>' -ForEach @(
        @{ Case = 'a registry clause without a key path'; Bad = @{ Type = 'RegistryKey'; RegistryKeyRelative = '' } }
        @{ Case = 'a registry value clause without a key path'; Bad = @{ Type = 'RegistryKeyValue'; RegistryKeyRelative = '' } }
        @{ Case = 'a file clause without a file name'; Bad = @{ Type = 'File'; FilePath = 'C:\Windows'; FileName = '' } }
        @{ Case = 'a script clause without a script'; Bad = @{ Type = 'Script'; ScriptText = '' } }
        @{ Case = 'a registry clause with no key property at all'; Bad = @{ Type = 'RegistryKey' } }
    ) {
        { Test-CompoundDetection -Connector Or -Clauses @($Bad, $script:Yes) } | Should -Not -Throw
        $res = Test-CompoundDetection -Connector Or -Clauses @($Bad, $script:Yes)
        $res.ClauseResults.Count | Should -Be 2
        $res.ClauseResults[0].Detected | Should -BeFalse
        $res.ClauseResults[0].Details | Should -Be 'The clause could not be evaluated. See the log for details.'
        $res.ClauseResults[1].Detected | Should -BeTrue
        $res.Detected | Should -BeTrue
        (Test-CompoundDetection -Connector And -Clauses @($Bad, $script:Yes)).Detected | Should -BeFalse
    }

    It 'rejects an empty list of clauses before it runs' {
        { Test-CompoundDetection -Connector And -Clauses @() } | Should -Throw
    }

    It 'reads Is64Bit "<Value>" as <Expected>' -ForEach @(
        @{ Value = $true;    Expected = $true }
        @{ Value = $false;   Expected = $false }
        @{ Value = 'false';  Expected = $false }
        @{ Value = 'true';   Expected = $true }
        @{ Value = 0;        Expected = $false }
    ) {
        # An HKLM ARP entry of a 64-bit application is visible only in the 64-bit view, so Is64Bit decides the result.
        $x64 = Get-ViewExclusiveX64App
        if (-not $x64) { Set-ItResult -Skipped -Because 'this machine has no 64-bit machine-wide application that the 32-bit view hides'; return }
        $clause = @{ Type = 'RegistryKey'; RegistryKeyRelative = $x64.RegistryKeyRelative; Hive = 'LocalMachine'; Is64Bit = $Value }
        (Test-CompoundDetection -Connector And -Clauses @($clause)).Detected | Should -Be $Expected
    }
}
