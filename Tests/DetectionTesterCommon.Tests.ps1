#Requires -Modules Pester

<#
.SYNOPSIS
    Pester tests for DetectionTesterCommon module.
#>

BeforeAll {
    $modulePath = Join-Path $PSScriptRoot "..\Module\DetectionTesterCommon.psd1"
    Import-Module $modulePath -Force -DisableNameChecking
}

Describe "Script parse validation" {
    It "start-detectiontester.ps1 parses without errors" {
        $scriptPath = Join-Path $PSScriptRoot "..\start-detectiontester.ps1"
        $tokens = $null
        $errors = $null
        [System.Management.Automation.Language.Parser]::ParseFile($scriptPath, [ref]$tokens, [ref]$errors) | Out-Null
        $errors.Count | Should -Be 0
    }

    It "DetectionTesterCommon.psm1 parses without errors" {
        $modulePath = Join-Path $PSScriptRoot "..\Module\DetectionTesterCommon.psm1"
        $tokens = $null
        $errors = $null
        [System.Management.Automation.Language.Parser]::ParseFile($modulePath, [ref]$tokens, [ref]$errors) | Out-Null
        $errors.Count | Should -Be 0
    }
}

Describe "Module load and function exports" {
    It "exports expected functions" {
        $expected = @(
            'Initialize-Logging', 'Write-Log',
            'Test-RegistryKeyValueDetection', 'Test-RegistryKeyDetection',
            'Test-FileDetection', 'Test-ScriptDetection', 'Test-CompoundDetection',
            'Get-InstalledApplications', 'Import-DetectionManifest',
            'Export-DetectionResultsCsv', 'Export-DetectionResultsHtml'
        )
        $exported = (Get-Command -Module DetectionTesterCommon).Name
        foreach ($fn in $expected) {
            $exported | Should -Contain $fn
        }
    }
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

Describe "Get-InstalledApplications" {
    It "returns at least one application" {
        $apps = Get-InstalledApplications
        $apps.Count | Should -BeGreaterThan 0
    }

    It "returns objects with expected properties" {
        $apps = Get-InstalledApplications
        $first = $apps[0]
        $first.PSObject.Properties.Name | Should -Contain 'DisplayName'
        $first.PSObject.Properties.Name | Should -Contain 'Publisher'
        $first.PSObject.Properties.Name | Should -Contain 'DisplayVersion'
        $first.PSObject.Properties.Name | Should -Contain 'Architecture'
        $first.PSObject.Properties.Name | Should -Contain 'RegistryKey'
    }
}

Describe "Import-DetectionManifest" {
    BeforeAll {
        $testManifestDir = Join-Path $TestDrive "manifests"
        New-Item -ItemType Directory -Path $testManifestDir -Force | Out-Null
    }

    It "imports RegistryKeyValue manifest" {
        $manifest = @{
            AppName = "Test App"
            Detection = @{
                Type = "RegistryKeyValue"
                RegistryKeyRelative = "SOFTWARE\Test\App"
                ValueName = "DisplayVersion"
                ExpectedValue = "1.0.0"
                Operator = "IsEquals"
                Is64Bit = $true
            }
        } | ConvertTo-Json -Depth 5
        $path = Join-Path $testManifestDir "regkeyval.json"
        Set-Content -LiteralPath $path -Value $manifest -Encoding UTF8

        $det = Import-DetectionManifest -Path $path
        $det.Type | Should -Be "RegistryKeyValue"
        $det.RegistryKeyRelative | Should -Be "SOFTWARE\Test\App"
        $det.ValueName | Should -Be "DisplayVersion"
        $det.ExpectedValue | Should -Be "1.0.0"
        $det.Operator | Should -Be "IsEquals"
        $det.Is64Bit | Should -Be $true
    }

    It "imports File manifest" {
        $manifest = @{
            AppName = "Test App 2"
            Detection = @{
                Type = "File"
                FilePath = "C:\Program Files\Test"
                FileName = "test.exe"
                PropertyType = "Version"
                Operator = "GreaterEquals"
                ExpectedValue = "2.0.0"
            }
        } | ConvertTo-Json -Depth 5
        $path = Join-Path $testManifestDir "file.json"
        Set-Content -LiteralPath $path -Value $manifest -Encoding UTF8

        $det = Import-DetectionManifest -Path $path
        $det.Type | Should -Be "File"
        $det.FilePath | Should -Be "C:\Program Files\Test"
        $det.FileName | Should -Be "test.exe"
        $det.PropertyType | Should -Be "Version"
        $det.Operator | Should -Be "GreaterEquals"
        $det.ExpectedValue | Should -Be "2.0.0"
    }

    It "normalizes DisplayVersion to ExpectedValue" {
        $manifest = @{
            AppName = "Normalized"
            Detection = @{
                Type = "RegistryKeyValue"
                RegistryKeyRelative = "SOFTWARE\Test"
                DisplayVersion = "3.0.0"
            }
        } | ConvertTo-Json -Depth 5
        $path = Join-Path $testManifestDir "normalized.json"
        Set-Content -LiteralPath $path -Value $manifest -Encoding UTF8

        $det = Import-DetectionManifest -Path $path
        $det.ExpectedValue | Should -Be "3.0.0"
        $det.ValueName | Should -Be "DisplayVersion"
        $det.Operator | Should -Be "IsEquals"
    }

    It "throws on missing Detection block" {
        $manifest = @{ AppName = "No Detection" } | ConvertTo-Json
        $path = Join-Path $testManifestDir "nodet.json"
        Set-Content -LiteralPath $path -Value $manifest -Encoding UTF8

        { Import-DetectionManifest -Path $path } | Should -Throw "*No 'Detection' block*"
    }

    It "throws on nonexistent file" {
        { Import-DetectionManifest -Path "C:\nonexistent\manifest.json" } | Should -Throw "*not found*"
    }
}
