#Requires -Modules Pester
# File detection.

BeforeAll {
    . "$PSScriptRoot\TestHelpers.ps1"
    Import-AppModule
    Mock Write-Log {} -ModuleName DetectionTesterCommon
}
Describe "Test-FileDetection" {
    It "detects existing file (explorer.exe existence)" {
        $result = Test-FileDetection `
            -FilePath "C:\Windows" `
            -FileName "explorer.exe" `
            -PropertyType "Existence"

        $result.Detected | Should -Be $true
        $result.FileExists | Should -Be $true
    }

    It "detects file version with GreaterEquals" {
        $result = Test-FileDetection `
            -FilePath "C:\Windows" `
            -FileName "explorer.exe" `
            -PropertyType "Version" `
            -ExpectedValue "1.0.0.0" `
            -Operator "GreaterEquals"

        $result.Detected | Should -Be $true
        $result.FileExists | Should -Be $true
        $result.ActualValue | Should -Not -BeNullOrEmpty
    }

    It "returns not detected for nonexistent file" {
        $result = Test-FileDetection `
            -FilePath "C:\Windows" `
            -FileName "nonexistent12345.exe" `
            -PropertyType "Existence"

        $result.Detected | Should -Be $false
        $result.FileExists | Should -Be $false
    }

    It "returns not detected when version is too high" {
        $result = Test-FileDetection `
            -FilePath "C:\Windows" `
            -FileName "explorer.exe" `
            -PropertyType "Version" `
            -ExpectedValue "999.0.0.0" `
            -Operator "GreaterEquals"

        $result.Detected | Should -Be $false
        $result.FileExists | Should -Be $true
    }
}

Describe 'Test-FileDetection on files in a test folder' {
    BeforeAll {
        $script:Dir = Join-Path $TestDrive 'files'
        New-Item -ItemType Directory -Path $script:Dir -Force | Out-Null
        Set-Content -LiteralPath (Join-Path $script:Dir 'plain.txt') -Value 'text'
        New-Item -ItemType File -Path (Join-Path $script:Dir 'empty.dat') | Out-Null
        Set-Content -LiteralPath (Join-Path $script:Dir 'with space.txt') -Value 'x'
        Set-Content -LiteralPath (Join-Path $script:Dir 'a[1].txt') -Value 'x'
        Set-Content -LiteralPath (Join-Path $script:Dir "Caf$([char]0xE9) $([char]0xFC).txt") -Value 'x'
        New-Item -ItemType Directory -Path (Join-Path $script:Dir 'folder.exe') | Out-Null
        New-Item -ItemType Directory -Path (Join-Path $script:Dir 'sub dir') | Out-Null
        Set-Content -LiteralPath (Join-Path $script:Dir 'sub dir\inner.txt') -Value 'x'
        $env:ZZTEST_DT_DIR = $script:Dir

        # A program with a known file version, built with the .NET Framework compiler that ships with Windows.
        $script:ExePath = $null
        $csc = Join-Path $env:windir 'Microsoft.NET\Framework64\v4.0.30319\csc.exe'
        if (Test-Path -LiteralPath $csc) {
            $src = Join-Path $script:Dir 'zz.cs'
            Set-Content -LiteralPath $src -Value '[assembly: System.Reflection.AssemblyFileVersion("3.4.5.6")] class P { static void Main() { } }'
            $exe = Join-Path $script:Dir 'versioned.exe'
            $null = & $csc /nologo /target:exe "/out:$exe" $src 2>&1
            if (Test-Path -LiteralPath $exe) { $script:ExePath = $exe }
        }
    }
    AfterAll { Remove-Item Env:\ZZTEST_DT_DIR -ErrorAction SilentlyContinue }

    It 'detects <Name> by existence' -ForEach @(
        @{ Name = 'a plain file';                 Dir = 'files';  File = 'plain.txt' }
        @{ Name = 'an empty file';                Dir = 'files';  File = 'empty.dat' }
        @{ Name = 'a file name with a space';     Dir = 'files';  File = 'with space.txt' }
        @{ Name = 'a file name with brackets';    Dir = 'files';  File = 'a[1].txt' }
        @{ Name = 'a file name in another case';  Dir = 'files';  File = 'PLAIN.TXT' }
        @{ Name = 'a file in a folder with a space'; Dir = 'files\sub dir'; File = 'inner.txt' }
    ) {
        $r = Test-FileDetection -FilePath (Join-Path $TestDrive $Dir) -FileName $File -PropertyType Existence
        $r.Detected | Should -BeTrue
        $r.FileExists | Should -BeTrue
        $r.Type | Should -Be 'File'
        $r.Operator | Should -Be 'Existence'
        $r.ExpectedValue | Should -Be '(file exists)'
        $r.Target | Should -Be (Join-Path (Join-Path $TestDrive $Dir) $File)
        $r.Details | Should -BeLike 'File exists: *'
    }

    It 'detects a file whose name has non-ASCII letters' {
        $name = "Caf$([char]0xE9) $([char]0xFC).txt"
        (Test-FileDetection -FilePath $script:Dir -FileName $name -PropertyType Existence).Detected | Should -BeTrue
    }

    It 'expands an environment variable in the folder' {
        $r = Test-FileDetection -FilePath '%ZZTEST_DT_DIR%' -FileName 'plain.txt' -PropertyType Existence
        $r.Detected | Should -BeTrue
        $r.Details | Should -BeLike ('File exists: ' + $script:Dir + '*')
    }

    It 'does not detect a file in a folder that uses an undefined environment variable' {
        $r = Test-FileDetection -FilePath '%ZZTEST_DT_UNDEFINED%\x' -FileName 'plain.txt' -PropertyType Existence
        $r.Detected | Should -BeFalse
        $r.FileExists | Should -BeFalse
    }

    It 'does not detect <Name>' -ForEach @(
        @{ Name = 'a missing file';             Path = 'files';       File = 'missing.txt' }
        @{ Name = 'a missing folder';           Path = 'nofolder';    File = 'plain.txt' }
        @{ Name = 'a file name used as a pattern'; Path = 'files';    File = 'plain.*' }
        @{ Name = 'a bracket pattern';          Path = 'files';       File = 'a[0-9].txt' }
    ) {
        $r = Test-FileDetection -FilePath (Join-Path $TestDrive $Path) -FileName $File -PropertyType Existence
        $r.Detected | Should -BeFalse
        $r.FileExists | Should -BeFalse
        $r.Details | Should -BeLike 'File not found: *'
    }

    It 'does not detect a folder that has the name of the file' {
        $r = Test-FileDetection -FilePath $script:Dir -FileName 'folder.exe' -PropertyType Existence
        $r.Detected | Should -BeFalse
        $r.FileExists | Should -BeFalse
    }

    It 'does not detect a path with characters that Windows does not allow' {
        $r = Test-FileDetection -FilePath 'C:\Win:dows|?' -FileName 'a*' -PropertyType Existence 2>$null
        $r.Detected | Should -BeFalse
        $r.FileExists | Should -BeFalse
    }

    It 'reads the file version and compares it with <Operator> <Expected>: <Result>' -ForEach @(
        @{ Operator = 'GreaterEquals'; Expected = '3.4.5.6'; Result = $true }
        @{ Operator = 'GreaterEquals'; Expected = '3.4.5.7'; Result = $false }
        @{ Operator = 'GreaterEquals'; Expected = '1.0.0.0'; Result = $true }
        @{ Operator = 'GreaterEquals'; Expected = '3.4';     Result = $true }
        @{ Operator = 'GreaterThan';   Expected = '3.4.5.5'; Result = $true }
        @{ Operator = 'GreaterThan';   Expected = '3.4.5.6'; Result = $false }
        @{ Operator = 'LessEquals';    Expected = '3.4.5.6'; Result = $true }
        @{ Operator = 'LessEquals';    Expected = '3.4.5.5'; Result = $false }
        @{ Operator = 'LessThan';      Expected = '3.4.5.7'; Result = $true }
        @{ Operator = 'LessThan';      Expected = '3.4.5.6'; Result = $false }
        @{ Operator = 'LessThan';      Expected = '3.10';    Result = $true }
        @{ Operator = 'IsEquals';      Expected = '3.4.5.6'; Result = $true }
        @{ Operator = 'IsEquals';      Expected = '3.4.5';   Result = $false }
        @{ Operator = 'GreaterEquals'; Expected = 'abc';     Result = $false }
        @{ Operator = 'GreaterEquals'; Expected = '';        Result = $false }
        @{ Operator = 'Between';       Expected = '3.4.5.6'; Result = $false }
    ) {
        if (-not $script:ExePath) { Set-ItResult -Skipped -Because 'the .NET Framework compiler is not available'; return }
        $r = Test-FileDetection -FilePath $script:Dir -FileName 'versioned.exe' -PropertyType Version -ExpectedValue $Expected -Operator $Operator
        $r.FileExists | Should -BeTrue
        $r.ActualValue | Should -Be '3.4.5.6'
        $r.Detected | Should -Be $Result
        $r.Operator | Should -Be $Operator
        $r.ExpectedValue | Should -Be $Expected
    }

    It 'says MATCH or NO MATCH with the version in the details' {
        if (-not $script:ExePath) { Set-ItResult -Skipped -Because 'the .NET Framework compiler is not available'; return }
        (Test-FileDetection -FilePath $script:Dir -FileName 'versioned.exe' -PropertyType Version -ExpectedValue '3.0' -Operator GreaterEquals).Details | Should -Be "MATCH (GreaterEquals): version '3.4.5.6' vs '3.0'"
        (Test-FileDetection -FilePath $script:Dir -FileName 'versioned.exe' -PropertyType Version -ExpectedValue '4.0' -Operator GreaterEquals).Details | Should -Be "NO MATCH (GreaterEquals): version '3.4.5.6' vs '4.0'"
    }

    It 'reports a file without version information' {
        $r = Test-FileDetection -FilePath $script:Dir -FileName 'plain.txt' -PropertyType Version -ExpectedValue '1.0' -Operator GreaterEquals
        $r.FileExists | Should -BeTrue
        $r.Detected | Should -BeFalse
        $r.ActualValue | Should -BeNullOrEmpty
        $r.Details | Should -Be 'File exists but has no version information'
    }

    It 'reports an unsupported property type' {
        $r = Test-FileDetection -FilePath $script:Dir -FileName 'plain.txt' -PropertyType Hash -ExpectedValue 'x'
        $r.FileExists | Should -BeTrue
        $r.Detected | Should -BeFalse
        $r.Details | Should -Be 'Unsupported PropertyType: Hash'
    }

    It 'ignores the expected value and the operator for an existence test' {
        $r = Test-FileDetection -FilePath $script:Dir -FileName 'plain.txt' -PropertyType Existence -ExpectedValue '9.9' -Operator LessThan
        $r.Detected | Should -BeTrue
        $r.Operator | Should -Be 'Existence'
        $r.ExpectedValue | Should -Be '(file exists)'
    }

    It 'rewrites System32 to SysWOW64 for a 32-bit test on a 64-bit system' {
        if (-not [Environment]::Is64BitOperatingSystem) { Set-ItResult -Skipped -Because 'this system is 32-bit'; return }
        $r = Test-FileDetection -FilePath (Join-Path $env:SystemRoot 'System32') -FileName 'notepad.exe' -PropertyType Existence -Is64Bit $false
        $r.Detected | Should -BeTrue
        $r.Details | Should -BeLike '*\SysWOW64\notepad.exe'
        $r64 = Test-FileDetection -FilePath (Join-Path $env:SystemRoot 'System32') -FileName 'notepad.exe' -PropertyType Existence -Is64Bit $true
        $r64.Details | Should -BeLike '*\System32\notepad.exe'
    }

    It 'rewrites System32 in the middle of a path, in any case, and nothing else' -ForEach @(
        @{ Path = 'C:\ZZ\system32\sub';   Shown = 'C:\ZZ\SysWOW64\sub\x.dll' }
        @{ Path = 'C:\ZZ\System32';       Shown = 'C:\ZZ\SysWOW64\x.dll' }
        @{ Path = 'C:\ZZ\System32Extra';  Shown = 'C:\ZZ\System32Extra\x.dll' }
        @{ Path = 'C:\ZZ\MySystem32';     Shown = 'C:\ZZ\MySystem32\x.dll' }
    ) {
        if (-not [Environment]::Is64BitOperatingSystem) { Set-ItResult -Skipped -Because 'this system is 32-bit'; return }
        $r = Test-FileDetection -FilePath $Path -FileName 'x.dll' -PropertyType Existence -Is64Bit $false
        $r.Details | Should -Be ('File not found: ' + $Shown)
    }
}

Describe 'Test-FileDetection for a 64-bit test' {
    It 'keeps the path when the test is 64-bit' {
        $r = Test-FileDetection -FilePath 'C:\ZZ\System32' -FileName 'x.dll' -PropertyType Existence -Is64Bit $true
        $r.Details | Should -Be 'File not found: C:\ZZ\System32\x.dll'
    }
}
