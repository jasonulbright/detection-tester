#Requires -Modules @{ ModuleName = 'Pester'; ModuleVersion = '5.1.0' }
# Value comparison (Compare-DetectionValue) and file version cleaning.

BeforeAll {
    . "$PSScriptRoot\TestHelpers.ps1"
    Import-AppModule
    Mock Write-Log {} -ModuleName DetectionTesterCommon
}

Describe 'Compare-DetectionValue IsEquals' {
    It 'compares "<Actual>" with "<Expected>" and gives <Result>' -ForEach @(
        @{ Actual = 'abc'; Expected = 'abc'; Result = $true }
        @{ Actual = 'ABC'; Expected = 'abc'; Result = $true }
        @{ Actual = 'abc'; Expected = 'abd'; Result = $false }
        @{ Actual = '';    Expected = '';    Result = $true }
        @{ Actual = '';    Expected = 'x';   Result = $false }
        @{ Actual = 'x';   Expected = '';    Result = $false }
        @{ Actual = '1.0'; Expected = '1.0.0'; Result = $false }
        @{ Actual = ' a';  Expected = 'a';   Result = $false }
        @{ Actual = '0';   Expected = '00';  Result = $false }
    ) {
        InModuleScope DetectionTesterCommon -Parameters @{ Actual = $Actual; Expected = $Expected; Result = $Result } {
            Compare-DetectionValue -Actual $Actual -Expected $Expected -Operator 'IsEquals' | Should -Be $Result
        }
    }

    # The parameters are typed [string], so PowerShell turns $null into ''. A caller never passes $null for a
    # value that exists: the detection functions stop earlier when the value is missing.
    It 'treats a null value as an empty string and does not throw' {
        InModuleScope DetectionTesterCommon {
            Compare-DetectionValue -Actual '' -Expected $null -Operator 'IsEquals' | Should -BeTrue
            Compare-DetectionValue -Actual $null -Expected '1' -Operator 'GreaterThan' -PropertyType 'Integer' | Should -BeFalse
            Compare-DetectionValue -Actual $null -Expected $null -Operator 'GreaterThan' -PropertyType 'Version' | Should -BeFalse
        }
    }
}

Describe 'Compare-DetectionValue ordering operators' {
    It 'compares <Actual> <Operator> <Expected> as <PropertyType> and gives <Result>' -ForEach @(
        # Integer kinds
        @{ PropertyType = 'Integer'; Operator = 'GreaterEquals'; Actual = '5';  Expected = '5';  Result = $true }
        @{ PropertyType = 'Integer'; Operator = 'GreaterEquals'; Actual = '4';  Expected = '5';  Result = $false }
        @{ PropertyType = 'Integer'; Operator = 'GreaterThan';   Actual = '5';  Expected = '5';  Result = $false }
        @{ PropertyType = 'Integer'; Operator = 'GreaterThan';   Actual = '10'; Expected = '9';  Result = $true }
        @{ PropertyType = 'Integer'; Operator = 'LessEquals';    Actual = '5';  Expected = '5';  Result = $true }
        @{ PropertyType = 'Integer'; Operator = 'LessEquals';    Actual = '6';  Expected = '5';  Result = $false }
        @{ PropertyType = 'Integer'; Operator = 'LessThan';      Actual = '5';  Expected = '5';  Result = $false }
        @{ PropertyType = 'Integer'; Operator = 'LessThan';      Actual = '-1'; Expected = '0';  Result = $true }
        @{ PropertyType = 'DWORD';   Operator = 'GreaterThan';   Actual = '4294967295'; Expected = '4294967294'; Result = $true }
        @{ PropertyType = 'QWORD';   Operator = 'GreaterThan';   Actual = '9223372036854775807'; Expected = '9223372036854775806'; Result = $true }
        @{ PropertyType = 'Int32';   Operator = 'LessThan';      Actual = '1';  Expected = '2';  Result = $true }
        @{ PropertyType = 'Int64';   Operator = 'LessThan';      Actual = '2';  Expected = '1';  Result = $false }
        @{ PropertyType = 'Int';     Operator = 'GreaterEquals'; Actual = ' 7 '; Expected = '7'; Result = $true }
        # Version kinds: components compare as numbers, not as text
        @{ PropertyType = 'Version'; Operator = 'GreaterThan';   Actual = '1.10';    Expected = '1.9';      Result = $true }
        @{ PropertyType = 'Version'; Operator = 'LessThan';      Actual = '1.9';     Expected = '1.10';     Result = $true }
        @{ PropertyType = 'Version'; Operator = 'GreaterEquals'; Actual = '2.0.0.0'; Expected = '2.0.0.0'; Result = $true }
        @{ PropertyType = 'Version'; Operator = 'LessEquals';    Actual = '2.0.0.1'; Expected = '2.0.0.0'; Result = $false }
        @{ PropertyType = 'FileVersion';    Operator = 'GreaterThan'; Actual = '10.0.19041.1'; Expected = '10.0.19041.0'; Result = $true }
        @{ PropertyType = 'ProductVersion'; Operator = 'LessThan';    Actual = '1.2.3';        Expected = '1.2.4';        Result = $true }
        # A single number is not a version; the integer path decides
        @{ PropertyType = 'Version'; Operator = 'GreaterThan';   Actual = '8246'; Expected = '99'; Result = $true }
        @{ PropertyType = 'Version'; Operator = 'LessThan';      Actual = '8246'; Expected = '99'; Result = $false }
        # Text that is neither a version nor a number never matches
        @{ PropertyType = 'Version'; Operator = 'GreaterEquals'; Actual = 'beta'; Expected = '1.0'; Result = $false }
        @{ PropertyType = 'Version'; Operator = 'GreaterEquals'; Actual = '1.0';  Expected = 'beta'; Result = $false }
        @{ PropertyType = 'Integer'; Operator = 'GreaterEquals'; Actual = '1.5';  Expected = '1';    Result = $false }
        @{ PropertyType = 'Integer'; Operator = 'GreaterThan';   Actual = '2';    Expected = '1.5';  Result = $false }
        @{ PropertyType = 'Integer'; Operator = 'LessThan';      Actual = '1.4';  Expected = '2';    Result = $false }
        @{ PropertyType = 'Integer'; Operator = 'GreaterEquals'; Actual = '';     Expected = '1';    Result = $false }
        @{ PropertyType = 'Integer'; Operator = 'GreaterEquals'; Actual = '99999999999999999999'; Expected = '1'; Result = $false }
        # Operators and kinds that do not exist never match
        @{ PropertyType = 'Integer'; Operator = 'NotEquals';     Actual = '1'; Expected = '2'; Result = $false }
        @{ PropertyType = 'Version'; Operator = 'Between';       Actual = '1.0'; Expected = '2.0'; Result = $false }
        @{ PropertyType = 'Boolean'; Operator = 'GreaterThan';   Actual = '1'; Expected = '0'; Result = $false }
        @{ PropertyType = 'Integer'; Operator = '';              Actual = '1'; Expected = '0'; Result = $false }
    ) {
        InModuleScope DetectionTesterCommon -Parameters @{ Actual = $Actual; Expected = $Expected; Operator = $Operator; PropertyType = $PropertyType; Result = $Result } {
            Compare-DetectionValue -Actual $Actual -Expected $Expected -Operator $Operator -PropertyType $PropertyType | Should -Be $Result
        }
    }

    # Decision: the shell sends PropertyType 'String' for every registry value, and the ordering operators
    # must work on a DisplayVersion such as 1.10 against 1.9. So 'String' orders by version, then by integer.
    # It never orders by text. The help text of Compare-DetectionValue says this.
    It 'orders the String kind by version and then by integer, never by text' -ForEach @(
        @{ Actual = '1.10'; Operator = 'GreaterThan'; Expected = '1.9';  Result = $true }
        @{ Actual = '1.9';  Operator = 'LessThan';    Expected = '1.10'; Result = $true }
        @{ Actual = '10';   Operator = 'GreaterThan'; Expected = '9';    Result = $true }
        @{ Actual = 'b';    Operator = 'GreaterThan'; Expected = 'a';    Result = $false }
        @{ Actual = 'a';    Operator = 'LessThan';    Expected = 'b';    Result = $false }
    ) {
        InModuleScope DetectionTesterCommon -Parameters @{ Actual = $Actual; Expected = $Expected; Operator = $Operator; Result = $Result } {
            Compare-DetectionValue -Actual $Actual -Expected $Expected -Operator $Operator -PropertyType 'String' | Should -Be $Result
        }
    }

    It 'documents the String ordering in its help text' {
        InModuleScope DetectionTesterCommon {
            $help = (Get-Help Compare-DetectionValue).description.Text -join "`n"
            $help | Should -Match "'String'.*version"
        }
    }

    It 'follows System.Version rules when the two versions have a different number of parts' {
        InModuleScope DetectionTesterCommon {
            Compare-DetectionValue -Actual '1.0.0' -Expected '1.0' -Operator 'GreaterEquals' -PropertyType 'Version' | Should -BeTrue
            Compare-DetectionValue -Actual '1.0'   -Expected '1.0.0' -Operator 'GreaterEquals' -PropertyType 'Version' | Should -BeFalse
        }
    }

    It 'logs a warning when a value cannot be parsed' {
        InModuleScope DetectionTesterCommon {
            Compare-DetectionValue -Actual 'x' -Expected '1' -Operator 'GreaterThan' -PropertyType 'Integer' | Should -BeFalse
        }
        Should -Invoke Write-Log -ModuleName DetectionTesterCommon -ParameterFilter { $Level -eq 'WARN' } -Times 1 -Exactly
    }
}

# A file version string from the version resource goes through ConvertTo-DTCleanFileVersion before it is compared.
Describe 'ConvertTo-DTCleanFileVersion' {
    It 'cleans "<Raw>" to "<Clean>"' -ForEach @(
        @{ Raw = '1.2.3.4';                                Clean = '1.2.3.4' }
        @{ Raw = '  1.2.3.4  ';                            Clean = '1.2.3.4' }
        @{ Raw = '10.0.19041.1 (WinBuild.160101.0800)';    Clean = '10.0.19041.1' }
        @{ Raw = '6.3.9600.17415 (winblue_r4.141028-1500)'; Clean = '6.3.9600.17415' }
        @{ Raw = '1, 0, 0, 1';                             Clean = '1.0.0.1' }
        @{ Raw = '1,0,0,1';                                Clean = '1.0.0.1' }
        @{ Raw = '2, 5';                                   Clean = '2.5' }
        @{ Raw = '1. 2. 3';                                Clean = '1.2.3' }
        @{ Raw = '3.4.5.6 built by me';                    Clean = '3.4.5.6' }
        @{ Raw = '1.0.0-beta';                             Clean = '1.0.0-beta' }
        @{ Raw = 'unknown';                                Clean = 'unknown' }
    ) {
        InModuleScope DetectionTesterCommon -Parameters @{ Raw = $Raw; Clean = $Clean } {
            ConvertTo-DTCleanFileVersion -Version $Raw | Should -Be $Clean
        }
    }

    It 'gives a string that System.Version accepts for the common version resource forms' -ForEach @(
        @{ Raw = '1, 0, 0, 1' }
        @{ Raw = '10.0.1 (WinBuild.160101.0800)' }
        @{ Raw = '4.7.3062.0 built by: NET472REL1' }
    ) {
        InModuleScope DetectionTesterCommon -Parameters @{ Raw = $Raw } {
            { [version](ConvertTo-DTCleanFileVersion -Version $Raw) } | Should -Not -Throw
        }
    }

    It 'returns an empty string for an empty or missing version' {
        InModuleScope DetectionTesterCommon {
            ConvertTo-DTCleanFileVersion -Version '' | Should -Be ''
            ConvertTo-DTCleanFileVersion -Version $null | Should -Be ''
            ConvertTo-DTCleanFileVersion -Version '   ' | Should -Be ''
        }
    }
}

Describe 'ConvertTo-DTBoolean' {
    It 'reads <Value> (<Kind>) as <Result>' -ForEach @(
        @{ Value = $true;    Kind = 'bool';   Result = $true }
        @{ Value = $false;   Kind = 'bool';   Result = $false }
        @{ Value = 'true';   Kind = 'string'; Result = $true }
        @{ Value = 'True';   Kind = 'string'; Result = $true }
        @{ Value = 'false';  Kind = 'string'; Result = $false }
        @{ Value = 'FALSE';  Kind = 'string'; Result = $false }
        @{ Value = '0';      Kind = 'string'; Result = $false }
        @{ Value = 'no';     Kind = 'string'; Result = $false }
        @{ Value = ' false '; Kind = 'string'; Result = $false }
        @{ Value = '1';      Kind = 'string'; Result = $true }
        @{ Value = 'yes';    Kind = 'string'; Result = $true }
        @{ Value = 0;        Kind = 'number'; Result = $false }
        @{ Value = 1;        Kind = 'number'; Result = $true }
    ) {
        InModuleScope DetectionTesterCommon -Parameters @{ Value = $Value; Result = $Result } {
            ConvertTo-DTBoolean -Value $Value -Default (-not $Result) | Should -Be $Result
        }
    }

    It 'returns the default for a missing, empty or unknown value' -ForEach @(
        @{ Value = $null }
        @{ Value = '' }
        @{ Value = 'maybe' }
        @{ Value = [datetime]'2026-01-01' }
        @{ Value = [guid]::Empty }
    ) {
        InModuleScope DetectionTesterCommon -Parameters @{ Value = $Value } {
            ConvertTo-DTBoolean -Value $Value -Default $true | Should -BeTrue
            ConvertTo-DTBoolean -Value $Value -Default $false | Should -BeFalse
        }
    }
}
