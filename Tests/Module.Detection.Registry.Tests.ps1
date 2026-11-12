#Requires -Modules @{ ModuleName = 'Pester'; ModuleVersion = '5.1.0' }
# Registry key and registry value detection.

BeforeAll {
    . "$PSScriptRoot\TestHelpers.ps1"
    Import-AppModule
    Mock Write-Log {} -ModuleName DetectionTesterCommon
}
Describe "Test-RegistryKeyValueDetection" {
    It "detects known registry value (CurrentVersion ProductName)" {
        $result = Test-RegistryKeyValueDetection `
            -RegistryKeyRelative "SOFTWARE\Microsoft\Windows NT\CurrentVersion" `
            -ValueName "ProductName" `
            -ExpectedValue "Windows" `
            -Operator "IsEquals" `
            -Is64Bit $true

        # ProductName won't exactly equal "Windows" -- it's "Windows 10 Pro" or similar
        $result.KeyExists | Should -Be $true
        $result.ValueFound | Should -Be $true
        $result.ActualValue | Should -Not -BeNullOrEmpty
    }

    It "matches with IsEquals when value matches exactly" {
        # Read actual value first, then test equality
        $view = [Microsoft.Win32.RegistryView]::Registry64
        $hklm = [Microsoft.Win32.RegistryKey]::OpenBaseKey([Microsoft.Win32.RegistryHive]::LocalMachine, $view)
        $key = $hklm.OpenSubKey("SOFTWARE\Microsoft\Windows NT\CurrentVersion")
        $actual = $key.GetValue("CurrentBuild")
        $key.Close(); $hklm.Close()

        $result = Test-RegistryKeyValueDetection `
            -RegistryKeyRelative "SOFTWARE\Microsoft\Windows NT\CurrentVersion" `
            -ValueName "CurrentBuild" `
            -ExpectedValue $actual `
            -Operator "IsEquals" `
            -Is64Bit $true

        $result.Detected | Should -Be $true
    }

    It "returns not detected for wrong expected value" {
        $result = Test-RegistryKeyValueDetection `
            -RegistryKeyRelative "SOFTWARE\Microsoft\Windows NT\CurrentVersion" `
            -ValueName "CurrentBuild" `
            -ExpectedValue "99999" `
            -Operator "IsEquals" `
            -Is64Bit $true

        $result.Detected | Should -Be $false
        $result.KeyExists | Should -Be $true
        $result.ValueFound | Should -Be $true
    }

    It "returns not detected for nonexistent key" {
        $result = Test-RegistryKeyValueDetection `
            -RegistryKeyRelative "SOFTWARE\NonExistentKey12345" `
            -ValueName "Test" `
            -ExpectedValue "test" `
            -Is64Bit $true

        $result.Detected | Should -Be $false
        $result.KeyExists | Should -Be $false
    }
}

Describe "Test-RegistryKeyDetection" {
    It "detects existing registry key" {
        $result = Test-RegistryKeyDetection `
            -RegistryKeyRelative "SOFTWARE\Microsoft\Windows NT\CurrentVersion" `
            -Is64Bit $true

        $result.Detected | Should -Be $true
        $result.KeyExists | Should -Be $true
    }

    It "returns not detected for nonexistent key" {
        $result = Test-RegistryKeyDetection `
            -RegistryKeyRelative "SOFTWARE\NonExistentKey12345" `
            -Is64Bit $true

        $result.Detected | Should -Be $false
        $result.KeyExists | Should -Be $false
    }
}

Describe "Phase 4: HKCU registry detection" {
    It "detects an existing HKCU key (Environment is per-user, always present)" {
        $result = Test-RegistryKeyDetection `
            -RegistryKeyRelative "Environment" `
            -Hive "CurrentUser"
        $result.Detected | Should -Be $true
        $result.KeyExists | Should -Be $true
    }

    It "result.Target carries HKCU prefix when Hive=CurrentUser" {
        $result = Test-RegistryKeyDetection `
            -RegistryKeyRelative "Environment" `
            -Hive "CurrentUser"
        $result.Target | Should -BeLike "HKCU\*"
    }

    It "result.Target carries HKLM prefix when Hive omitted (default backward compat)" {
        $result = Test-RegistryKeyDetection `
            -RegistryKeyRelative "SOFTWARE\Microsoft\Windows NT\CurrentVersion"
        $result.Target | Should -BeLike "HKLM\*"
    }

    It "result.Hive field reflects the requested hive" {
        $r1 = Test-RegistryKeyDetection -RegistryKeyRelative "Environment" -Hive "CurrentUser"
        $r1.Hive | Should -Be "CurrentUser"

        $r2 = Test-RegistryKeyDetection -RegistryKeyRelative "SOFTWARE\Microsoft\Windows NT\CurrentVersion"
        $r2.Hive | Should -Be "LocalMachine"
    }

    It "RegistryKeyValue: HKCU value lookup works (TEMP env var)" {
        $result = Test-RegistryKeyValueDetection `
            -RegistryKeyRelative "Environment" `
            -ValueName "TEMP" `
            -ExpectedValue "ignored" `
            -Operator "IsEquals" `
            -Hive "CurrentUser"
        $result.KeyExists | Should -Be $true
        $result.ValueFound | Should -Be $true
        $result.ActualValue | Should -Not -BeNullOrEmpty
    }

    It "Hive parameter rejects values outside the validate set" {
        { Test-RegistryKeyDetection -RegistryKeyRelative "Environment" -Hive "BogusHive" } |
            Should -Throw
    }

    It "Compound detection forwards Hive to clause evaluators" {
        $clauses = @(
            @{ Type = 'RegistryKey'; RegistryKeyRelative = 'Environment'; Hive = 'CurrentUser' }
            @{ Type = 'RegistryKey'; RegistryKeyRelative = 'SOFTWARE\Microsoft\Windows NT\CurrentVersion'; Hive = 'LocalMachine' }
        )
        $result = Test-CompoundDetection -Connector 'And' -Clauses $clauses
        $result.Detected | Should -Be $true
        $result.ClauseResults[0].Hive | Should -Be "CurrentUser"
        $result.ClauseResults[1].Hive | Should -Be "LocalMachine"
    }
}

# The tests below read a throwaway key under HKCU\Software\ZZTEST-DT-<guid>, so every expected value is known.
Describe 'Test-RegistryKeyValueDetection on a throwaway key' {
    BeforeAll {
        $script:Key = New-TestRegistryKey -Value @{
            Str      = '1.2.3'
            Mixed    = 'Hello World'
            Unicode  = "Caf$([char]0xE9) $([char]0x00FC)"
            Empty    = ''
            Dw5      = @{ Data = 5; Kind = 'DWord' }
            Dw0      = @{ Data = 0; Kind = 'DWord' }
            DwMax    = @{ Data = [int]-1; Kind = 'DWord' }
            Qw       = @{ Data = [int64]4294967296; Kind = 'QWord' }
            Multi    = @{ Data = [string[]]@('a', 'b'); Kind = 'MultiString' }
            Bin      = @{ Data = [byte[]]@(1, 2, 255); Kind = 'Binary' }
            Expand   = @{ Data = '%SystemRoot%\ZZ'; Kind = 'ExpandString' }
            'Name With Space' = 'x'
            '[wild]*' = 'literal'
        }
    }
    AfterAll { Remove-TestRegistryKey -Relative $script:Key }

    It 'detects a string value that equals the expected value' {
        $r = Test-RegistryKeyValueDetection -RegistryKeyRelative $script:Key -ValueName 'Str' -ExpectedValue '1.2.3' -Hive CurrentUser
        $r.Detected | Should -BeTrue
        $r.KeyExists | Should -BeTrue
        $r.ValueFound | Should -BeTrue
        $r.ActualValue | Should -Be '1.2.3'
        $r.Details | Should -BeLike 'MATCH (IsEquals)*'
        $r.Type | Should -Be 'RegistryKeyValue'
        $r.Target | Should -Be ('HKCU\' + $script:Key)
        $r.Hive | Should -Be 'CurrentUser'
        $r.Operator | Should -Be 'IsEquals'
        $r.ExpectedValue | Should -Be '1.2.3'
    }

    It 'compares an equal string without regard to case' {
        (Test-RegistryKeyValueDetection -RegistryKeyRelative $script:Key -ValueName 'Mixed' -ExpectedValue 'hello world' -Hive CurrentUser).Detected | Should -BeTrue
    }

    It 'does not detect a different string and says NO MATCH' {
        $r = Test-RegistryKeyValueDetection -RegistryKeyRelative $script:Key -ValueName 'Mixed' -ExpectedValue 'Hello' -Hive CurrentUser
        $r.Detected | Should -BeFalse
        $r.ValueFound | Should -BeTrue
        $r.Details | Should -BeLike 'NO MATCH (IsEquals)*'
    }

    It 'matches an empty value to an empty expected value' {
        $r = Test-RegistryKeyValueDetection -RegistryKeyRelative $script:Key -ValueName 'Empty' -ExpectedValue '' -Hive CurrentUser
        $r.ValueFound | Should -BeTrue
        $r.Detected | Should -BeTrue
    }

    It 'keeps non-ASCII text intact' {
        $r = Test-RegistryKeyValueDetection -RegistryKeyRelative $script:Key -ValueName 'Unicode' -ExpectedValue "Caf$([char]0xE9) $([char]0x00FC)" -Hive CurrentUser
        $r.Detected | Should -BeTrue
    }

    It 'finds a value name that has a space and one that looks like a wildcard' -ForEach @(
        @{ Name = 'Name With Space'; Expected = 'x' }
        @{ Name = '[wild]*'; Expected = 'literal' }
    ) {
        $r = Test-RegistryKeyValueDetection -RegistryKeyRelative $script:Key -ValueName $Name -ExpectedValue $Expected -Hive CurrentUser
        $r.ValueFound | Should -BeTrue
        $r.Detected | Should -BeTrue
    }

    It 'does not treat a wildcard value name as a pattern' {
        $r = Test-RegistryKeyValueDetection -RegistryKeyRelative $script:Key -ValueName 'S*' -ExpectedValue '1.2.3' -Hive CurrentUser
        $r.ValueFound | Should -BeFalse
        $r.Detected | Should -BeFalse
    }

    It 'finds a value name in any letter case' {
        (Test-RegistryKeyValueDetection -RegistryKeyRelative $script:Key -ValueName 'sTR' -ExpectedValue '1.2.3' -Hive CurrentUser).Detected | Should -BeTrue
    }

    It 'compares a DWORD with <Operator> <Expected> and gives <Result>' -ForEach @(
        @{ Operator = 'IsEquals';      Expected = '5'; Result = $true }
        @{ Operator = 'IsEquals';      Expected = '6'; Result = $false }
        @{ Operator = 'GreaterEquals'; Expected = '5'; Result = $true }
        @{ Operator = 'GreaterEquals'; Expected = '6'; Result = $false }
        @{ Operator = 'GreaterThan';   Expected = '4'; Result = $true }
        @{ Operator = 'GreaterThan';   Expected = '5'; Result = $false }
        @{ Operator = 'LessEquals';    Expected = '5'; Result = $true }
        @{ Operator = 'LessEquals';    Expected = '4'; Result = $false }
        @{ Operator = 'LessThan';      Expected = '6'; Result = $true }
        @{ Operator = 'LessThan';      Expected = '5'; Result = $false }
    ) {
        $r = Test-RegistryKeyValueDetection -RegistryKeyRelative $script:Key -ValueName 'Dw5' -ExpectedValue $Expected -Operator $Operator -PropertyType 'DWORD' -Hive CurrentUser
        $r.Detected | Should -Be $Result
        $r.ActualValue | Should -Be '5'
    }

    It 'matches a DWORD of zero to an expected value of 0' {
        $r = Test-RegistryKeyValueDetection -RegistryKeyRelative $script:Key -ValueName 'Dw0' -ExpectedValue '0' -Hive CurrentUser
        $r.ActualValue | Should -Be '0'
        $r.Detected | Should -BeTrue
    }

    It 'reads a DWORD of 0xFFFFFFFF as a signed number' {
        (Test-RegistryKeyValueDetection -RegistryKeyRelative $script:Key -ValueName 'DwMax' -ExpectedValue '-1' -Hive CurrentUser).Detected | Should -BeTrue
    }

    It 'compares a QWORD above the Int32 range' {
        (Test-RegistryKeyValueDetection -RegistryKeyRelative $script:Key -ValueName 'Qw' -ExpectedValue '4294967295' -Operator GreaterThan -PropertyType QWORD -Hive CurrentUser).Detected | Should -BeTrue
        (Test-RegistryKeyValueDetection -RegistryKeyRelative $script:Key -ValueName 'Qw' -ExpectedValue '4294967297' -Operator LessThan -PropertyType QWORD -Hive CurrentUser).Detected | Should -BeTrue
    }

    It 'compares a version string by components for <PropertyType>' -ForEach @(
        @{ PropertyType = 'Version' }
        @{ PropertyType = 'String' }
    ) {
        $p = @{ RegistryKeyRelative = $script:Key; ValueName = 'Str'; PropertyType = $PropertyType; Hive = 'CurrentUser' }
        (Test-RegistryKeyValueDetection @p -ExpectedValue '1.2.3' -Operator GreaterEquals).Detected | Should -BeTrue
        (Test-RegistryKeyValueDetection @p -ExpectedValue '1.2.3' -Operator GreaterThan).Detected | Should -BeFalse
        (Test-RegistryKeyValueDetection @p -ExpectedValue '1.2.2' -Operator GreaterThan).Detected | Should -BeTrue
        (Test-RegistryKeyValueDetection @p -ExpectedValue '1.10' -Operator LessThan).Detected | Should -BeTrue
    }

    It 'does not throw and does not match when a numeric comparison meets text' {
        $r = Test-RegistryKeyValueDetection -RegistryKeyRelative $script:Key -ValueName 'Mixed' -ExpectedValue '5' -Operator GreaterThan -PropertyType DWORD -Hive CurrentUser
        $r.ValueFound | Should -BeTrue
        $r.Detected | Should -BeFalse
        $r.Details | Should -BeLike 'NO MATCH*'
    }

    It 'shows a multi-string value as its lines joined by a space' {
        $r = Test-RegistryKeyValueDetection -RegistryKeyRelative $script:Key -ValueName 'Multi' -ExpectedValue 'a b' -Hive CurrentUser
        $r.ActualValue | Should -Be 'a b'
        $r.ActualValue | Should -Not -BeLike '*System.*'
        $r.Detected | Should -BeTrue
    }

    It 'shows a binary value as its bytes in decimal, separated by a space' {
        $r = Test-RegistryKeyValueDetection -RegistryKeyRelative $script:Key -ValueName 'Bin' -ExpectedValue '1 2 255' -Hive CurrentUser
        $r.ActualValue | Should -Be '1 2 255'
        $r.ActualValue | Should -Not -BeLike '*System.*'
        $r.Detected | Should -BeTrue
    }

    It 'expands environment variables in an expandable string' {
        $r = Test-RegistryKeyValueDetection -RegistryKeyRelative $script:Key -ValueName 'Expand' -ExpectedValue ($env:SystemRoot + '\ZZ') -Hive CurrentUser
        $r.Detected | Should -BeTrue
    }

    It 'reports a value that is missing from an existing key' {
        $r = Test-RegistryKeyValueDetection -RegistryKeyRelative $script:Key -ValueName 'Nope' -ExpectedValue 'x' -Hive CurrentUser
        $r.KeyExists | Should -BeTrue
        $r.ValueFound | Should -BeFalse
        $r.Detected | Should -BeFalse
        $r.ActualValue | Should -BeNullOrEmpty
        $r.Details | Should -Be "Value 'Nope' not found in key"
    }

    It 'reports a key that is missing' {
        $r = Test-RegistryKeyValueDetection -RegistryKeyRelative ($script:Key + '\missing') -ValueName 'Str' -ExpectedValue 'x' -Hive CurrentUser
        $r.KeyExists | Should -BeFalse
        $r.ValueFound | Should -BeFalse
        $r.Detected | Should -BeFalse
        $r.Details | Should -BeLike 'Key not found: HKCU\*'
    }

    It 'finds the key in the 64-bit and the 32-bit view of HKCU' -ForEach @(@{ Is64 = $true }, @{ Is64 = $false }) {
        (Test-RegistryKeyValueDetection -RegistryKeyRelative $script:Key -ValueName 'Str' -ExpectedValue '1.2.3' -Is64Bit $Is64 -Hive CurrentUser).Detected | Should -BeTrue
    }

    It 'does not find the key under HKLM' {
        $r = Test-RegistryKeyValueDetection -RegistryKeyRelative $script:Key -ValueName 'Str' -ExpectedValue '1.2.3'
        $r.KeyExists | Should -BeFalse
        $r.Target | Should -BeLike 'HKLM\*'
        $r.Hive | Should -Be 'LocalMachine'
    }

    It 'returns an error result for a key path longer than the registry allows' {
        $r = Test-RegistryKeyValueDetection -RegistryKeyRelative ('Software\' + ('a' * 300)) -ValueName 'Str' -ExpectedValue '1' -Hive CurrentUser
        $r.Detected | Should -BeFalse
        $r.Details | Should -Be 'The registry value could not be read. See the log for details.'
    }

    It 'sees the 64-bit and the 32-bit view of HKLM differ for a 64-bit ARP entry' {
        $x64 = Get-ViewExclusiveX64App
        if (-not $x64) { Set-ItResult -Skipped -Because 'this machine has no 64-bit machine-wide application that the 32-bit view hides'; return }
        (Test-RegistryKeyDetection -RegistryKeyRelative $x64.RegistryKeyRelative -Is64Bit $true).Detected | Should -BeTrue
        (Test-RegistryKeyDetection -RegistryKeyRelative $x64.RegistryKeyRelative -Is64Bit $false).Detected | Should -BeFalse
    }
}

Describe 'Test-RegistryKeyDetection on a throwaway key' {
    BeforeAll {
        $script:Key = New-TestRegistryKey
        $sub = [Microsoft.Win32.Registry]::CurrentUser.CreateSubKey($script:Key + '\Child')
        $sub.Close()
    }
    AfterAll { Remove-TestRegistryKey -Relative $script:Key }

    It 'detects the key and a child key and reports the existence fields' {
        $r = Test-RegistryKeyDetection -RegistryKeyRelative $script:Key -Hive CurrentUser
        $r.Detected | Should -BeTrue
        $r.KeyExists | Should -BeTrue
        $r.Type | Should -Be 'RegistryKey'
        $r.Operator | Should -Be 'Existence'
        $r.ExpectedValue | Should -Be '(key exists)'
        $r.Target | Should -Be ('HKCU\' + $script:Key)
        $r.Details | Should -BeLike 'Key exists: HKCU\*'
        (Test-RegistryKeyDetection -RegistryKeyRelative ($script:Key + '\Child') -Hive CurrentUser).Detected | Should -BeTrue
    }

    It 'does not detect a missing child, and a key name with a different spelling is still the same key' {
        (Test-RegistryKeyDetection -RegistryKeyRelative ($script:Key + '\Other') -Hive CurrentUser).Detected | Should -BeFalse
        (Test-RegistryKeyDetection -RegistryKeyRelative ($script:Key + '\child') -Hive CurrentUser).Detected | Should -BeTrue
    }

    It 'treats a trailing backslash as the same key' {
        (Test-RegistryKeyDetection -RegistryKeyRelative ($script:Key + '\') -Hive CurrentUser).Detected | Should -BeTrue
    }

    It 'does not treat a wildcard as a pattern' {
        (Test-RegistryKeyDetection -RegistryKeyRelative ($script:Key + '\Ch*') -Hive CurrentUser).Detected | Should -BeFalse
    }

    It 'returns an error result for a key path longer than the registry allows' {
        $r = Test-RegistryKeyDetection -RegistryKeyRelative ('Software\' + ('a' * 300)) -Hive CurrentUser
        $r.Detected | Should -BeFalse
        $r.Details | Should -Be 'The registry key could not be read. See the log for details.'
    }

    It 'rejects an empty key path before it runs' {
        { Test-RegistryKeyDetection -RegistryKeyRelative '' } | Should -Throw
    }
}

Describe 'Registry handle cleanup' {
    # An exception between OpenBaseKey and Close leaves the base key open. This syntax check requires
    # that every function that opens a base key closes it in a finally block.
    It 'closes the base key in a finally block in <Name>' -ForEach @(
        @{ Name = 'Test-RegistryKeyValueDetection' }
        @{ Name = 'Test-RegistryKeyDetection' }
        @{ Name = 'Get-InstalledApplications' }
    ) {
        $ast = [System.Management.Automation.Language.Parser]::ParseFile((Join-Path $script:AppRoot 'Module\DetectionTesterCommon.psm1'), [ref]$null, [ref]$null)
        $fn = $ast.Find({ param($n) $n -is [System.Management.Automation.Language.FunctionDefinitionAst] -and $n.Name -eq $Name }, $true)
        $fn | Should -Not -BeNullOrEmpty
        $opens = @($fn.FindAll({ param($n) $n -is [System.Management.Automation.Language.InvokeMemberExpressionAst] -and $n.Member.Value -eq 'OpenBaseKey' }, $true))
        $opens.Count | Should -BeGreaterThan 0
        $finallies = @($fn.FindAll({ param($n) $n -is [System.Management.Automation.Language.TryStatementAst] -and $n.Finally }, $true) | ForEach-Object { $_.Finally })
        $finallies.Count | Should -BeGreaterThan 0
        $closed = @($finallies | ForEach-Object { $_.FindAll({ param($n) $n -is [System.Management.Automation.Language.InvokeMemberExpressionAst] -and $n.Member.Value -eq 'Close' }, $true) } | ForEach-Object { $_.Expression.Extent.Text })
        $closed | Should -Contain '$rootKey'
        # No Close call on the base key outside a finally block
        $all = @($fn.FindAll({ param($n) $n -is [System.Management.Automation.Language.InvokeMemberExpressionAst] -and $n.Member.Value -eq 'Close' -and $n.Expression.Extent.Text -eq '$rootKey' }, $true))
        $inFinally = @($finallies | ForEach-Object { $_.FindAll({ param($n) $n -is [System.Management.Automation.Language.InvokeMemberExpressionAst] -and $n.Member.Value -eq 'Close' -and $n.Expression.Extent.Text -eq '$rootKey' }, $true) })
        # Each acquisition is followed at once by a try whose finally closes the key.
        foreach ($open in $opens) {
            $stmt = $open
            while ($stmt.Parent -isnot [System.Management.Automation.Language.StatementBlockAst] -and $stmt.Parent -isnot [System.Management.Automation.Language.NamedBlockAst]) { $stmt = $stmt.Parent }
            $siblings = @($stmt.Parent.Statements)
            $next = $siblings[[array]::IndexOf($siblings, $stmt) + 1]
            $next | Should -BeOfType ([System.Management.Automation.Language.TryStatementAst]) -Because 'a statement between OpenBaseKey and the try can throw and leak the key'
            $next.Finally.Extent.Text | Should -Match '\$rootKey\.Close\(\)'
        }
        $all.Count | Should -Be $inFinally.Count -Because 'an early Close outside the finally block is skipped when an exception happens first'
    }
}

