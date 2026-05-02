#Requires -Modules @{ ModuleName = 'Pester'; ModuleVersion = '5.0' }

<#
.SYNOPSIS
    Pester tests for DetectionTesterCommon module.
#>

$modulePath = Join-Path $PSScriptRoot "..\Module\DetectionTesterCommon.psd1"
Import-Module $modulePath -Force -DisableNameChecking

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
        $first.PSObject.Properties.Name | Should -Contain 'RegistryKeyRelative'
        $first.PSObject.Properties.Name | Should -Contain 'Scope'
        $first.PSObject.Properties.Name | Should -Contain 'Hive'
    }

    It "Scope values are constrained to Machine or User" {
        $apps = Get-InstalledApplications
        $scopes = $apps.Scope | Select-Object -Unique
        foreach ($s in $scopes) {
            $s | Should -BeIn @('Machine', 'User')
        }
    }

    It "RegistryKey display string carries hive prefix" {
        $apps = Get-InstalledApplications
        $first = $apps[0]
        $first.RegistryKey | Should -Match '^HK(LM|CU)\\'
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

Describe "Codex review fixes" {
    Context "Compare-DetectionValue integer comparisons" {
        It "compares DWORD style values numerically (UBR style)" {
            InModuleScope DetectionTesterCommon {
                Compare-DetectionValue -Actual '8246' -Expected '1'     -Operator 'GreaterEquals' -PropertyType 'Integer' | Should -Be $true
                Compare-DetectionValue -Actual '8246' -Expected '99999' -Operator 'GreaterEquals' -PropertyType 'Integer' | Should -Be $false
            }
        }
    }

    Context "ConvertTo-DTHiveName alias normalization" {
        It "maps HKLM aliases to LocalMachine" {
            InModuleScope DetectionTesterCommon {
                ConvertTo-DTHiveName 'HKLM'                | Should -Be 'LocalMachine'
                ConvertTo-DTHiveName 'HKEY_LOCAL_MACHINE' | Should -Be 'LocalMachine'
                ConvertTo-DTHiveName 'LocalMachine'       | Should -Be 'LocalMachine'
            }
        }
        It "maps HKCU aliases to CurrentUser" {
            InModuleScope DetectionTesterCommon {
                ConvertTo-DTHiveName 'HKCU'                | Should -Be 'CurrentUser'
                ConvertTo-DTHiveName 'HKEY_CURRENT_USER'  | Should -Be 'CurrentUser'
                ConvertTo-DTHiveName 'CurrentUser'        | Should -Be 'CurrentUser'
            }
        }
    }

    Context "ConvertTo-DTRegistryRelativePath prefix stripping" {
        It "strips HKLM prefix" {
            InModuleScope DetectionTesterCommon {
                ConvertTo-DTRegistryRelativePath 'HKLM\SOFTWARE\Vendor\App' | Should -Be 'SOFTWARE\Vendor\App'
            }
        }
        It "strips HKEY_CURRENT_USER prefix" {
            InModuleScope DetectionTesterCommon {
                ConvertTo-DTRegistryRelativePath 'HKEY_CURRENT_USER\Software\Vendor' | Should -Be 'Software\Vendor'
            }
        }
        It "strips HKCU PSDrive colon form (HKCU:\...)" {
            InModuleScope DetectionTesterCommon {
                ConvertTo-DTRegistryRelativePath 'HKCU:\Software\Vendor' | Should -Be 'Software\Vendor'
            }
        }
        It "strips Registry:: provider prefix with HKEY_*" {
            InModuleScope DetectionTesterCommon {
                ConvertTo-DTRegistryRelativePath 'Registry::HKEY_CURRENT_USER\Software\Vendor' | Should -Be 'Software\Vendor'
            }
        }
        It "strips Microsoft.PowerShell.Core\Registry:: provider prefix" {
            InModuleScope DetectionTesterCommon {
                ConvertTo-DTRegistryRelativePath 'Microsoft.PowerShell.Core\Registry::HKLM\SOFTWARE\Vendor' | Should -Be 'SOFTWARE\Vendor'
            }
        }
        It "passes through paths with no prefix" {
            InModuleScope DetectionTesterCommon {
                ConvertTo-DTRegistryRelativePath 'SOFTWARE\Vendor\App' | Should -Be 'SOFTWARE\Vendor\App'
            }
        }
    }

    Context "Get-DTHiveFromPath provider-style detection" {
        It "detects HKCU from PSDrive form" {
            InModuleScope DetectionTesterCommon {
                Get-DTHiveFromPath 'HKCU:\Software\Vendor' | Should -Be 'CurrentUser'
            }
        }
        It "detects HKCU from Registry:: provider" {
            InModuleScope DetectionTesterCommon {
                Get-DTHiveFromPath 'Registry::HKEY_CURRENT_USER\Software\Vendor' | Should -Be 'CurrentUser'
            }
        }
        It "detects HKLM from Microsoft.PowerShell.Core provider" {
            InModuleScope DetectionTesterCommon {
                Get-DTHiveFromPath 'Microsoft.PowerShell.Core\Registry::HKLM\SOFTWARE\Vendor' | Should -Be 'LocalMachine'
            }
        }
        It "returns empty for paths with no hive prefix" {
            InModuleScope DetectionTesterCommon {
                Get-DTHiveFromPath 'SOFTWARE\Vendor\App' | Should -Be ''
            }
        }
    }

    Context "Manifest import: provider-style paths are normalized" {
        BeforeAll {
            $tmp = Join-Path $TestDrive "provider-paths"
            New-Item -ItemType Directory -Path $tmp -Force | Out-Null
        }
        It "imports HKCU:\... PSDrive form as CurrentUser" {
            $manifest = @{
                Detection = @{
                    Type = 'RegistryKeyValue'
                    KeyPath = 'HKCU:\Software\Vendor\App'
                    ValueName = 'DisplayVersion'
                    ExpectedValue = '1.0.0'
                }
            } | ConvertTo-Json -Depth 5
            $p = Join-Path $tmp "psdrive.json"
            Set-Content -LiteralPath $p -Value $manifest -Encoding UTF8
            $det = Import-DetectionManifest -Path $p
            $det.Hive                | Should -Be 'CurrentUser'
            $det.RegistryKeyRelative | Should -Be 'Software\Vendor\App'
        }
        It "imports Registry::HKEY_CURRENT_USER\... as CurrentUser" {
            $manifest = @{
                Detection = @{
                    Type = 'RegistryKey'
                    RegistryKeyRelative = 'Registry::HKEY_CURRENT_USER\Software\Vendor'
                }
            } | ConvertTo-Json -Depth 5
            $p = Join-Path $tmp "registry-provider.json"
            Set-Content -LiteralPath $p -Value $manifest -Encoding UTF8
            $det = Import-DetectionManifest -Path $p
            $det.Hive                | Should -Be 'CurrentUser'
            $det.RegistryKeyRelative | Should -Be 'Software\Vendor'
        }
    }

    Context "HTML report title escaping" {
        It "encodes the ReportTitle parameter" {
            $dt = New-Object System.Data.DataTable
            foreach ($c in @('Time','Type','Mode','Target','Expected','Found','Result')) {
                [void]$dt.Columns.Add($c, [string])
            }
            $row = $dt.NewRow()
            $row['Time']='00:00:00'; $row['Type']='Script'; $row['Mode']='Positive'
            $row['Target']='x'; $row['Expected']='y'; $row['Found']='z'; $row['Result']='Pass'
            [void]$dt.Rows.Add($row)
            $out = Join-Path $TestDrive "title-escape.html"
            $injected = '<img src=x onerror=alert(1)>'
            Export-DetectionResultsHtml -DataTable $dt -OutputPath $out -ReportTitle $injected
            $html = Get-Content -LiteralPath $out -Raw
            $html | Should -Not -Match '<img src=x onerror=alert'
            $html | Should -Match '&lt;img'
        }
    }

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

    Context "Manifest path + hive normalization" {
        BeforeAll {
            $tmp = Join-Path $TestDrive "codex-manifests"
            New-Item -ItemType Directory -Path $tmp -Force | Out-Null
        }
        It "imports a manifest with HKCU\\... full path" {
            $manifest = @{
                Detection = @{
                    Type = 'RegistryKeyValue'
                    KeyPath = 'HKCU\Software\Vendor\App'
                    ValueName = 'DisplayVersion'
                    ExpectedValue = '1.0.0'
                }
            } | ConvertTo-Json -Depth 5
            $p = Join-Path $tmp "hkcu-fullpath.json"
            Set-Content -LiteralPath $p -Value $manifest -Encoding UTF8
            $det = Import-DetectionManifest -Path $p
            $det.Hive                | Should -Be 'CurrentUser'
            $det.RegistryKeyRelative | Should -Be 'Software\Vendor\App'
        }
        It "normalizes HKLM alias in Hive field" {
            $manifest = @{
                Detection = @{
                    Type = 'RegistryKey'
                    RegistryKeyRelative = 'SOFTWARE\Vendor'
                    Hive = 'HKLM'
                }
            } | ConvertTo-Json -Depth 5
            $p = Join-Path $tmp "hkcu-alias.json"
            Set-Content -LiteralPath $p -Value $manifest -Encoding UTF8
            $det = Import-DetectionManifest -Path $p
            $det.Hive | Should -Be 'LocalMachine'
        }
    }

    Context "HTML export escaping" {
        It "encodes special characters in cell values" {
            $dt = New-Object System.Data.DataTable
            foreach ($c in @('Time','Type','Mode','Target','Expected','Found','Result')) {
                [void]$dt.Columns.Add($c, [string])
            }
            $row = $dt.NewRow()
            $row['Time']='00:00:00'; $row['Type']='Script'; $row['Mode']='Positive'
            $row['Target']='<img src=x onerror=alert(1)>'
            $row['Expected']='a & b'
            $row['Found']='1 < 2'
            $row['Result']='Pass'
            [void]$dt.Rows.Add($row)
            $out = Join-Path $TestDrive "html-escape.html"
            Export-DetectionResultsHtml -DataTable $dt -OutputPath $out
            $html = Get-Content -LiteralPath $out -Raw
            $html | Should -Not -Match '<img src=x onerror=alert'
            $html | Should -Match '&lt;img'
            $html | Should -Match 'a &amp; b'
            $html | Should -Match '1 &lt; 2'
        }
    }
}

Describe "Phase 4: Manifest Hive normalization" {
    BeforeAll {
        $testManifestDir = Join-Path $TestDrive "manifests-hkcu"
        New-Item -ItemType Directory -Path $testManifestDir -Force | Out-Null
    }

    It "imports RegistryKeyValue manifest with explicit Hive=CurrentUser" {
        $manifest = @{
            Detection = @{
                Type = "RegistryKeyValue"
                RegistryKeyRelative = "Software\Vendor\App"
                ValueName = "DisplayVersion"
                ExpectedValue = "1.0.0"
                Hive = "CurrentUser"
            }
        } | ConvertTo-Json -Depth 5
        $path = Join-Path $testManifestDir "hkcu-rkv.json"
        Set-Content -LiteralPath $path -Value $manifest -Encoding UTF8

        $det = Import-DetectionManifest -Path $path
        $det.Hive | Should -Be "CurrentUser"
    }

    It "defaults Hive to LocalMachine when manifest omits it" {
        $manifest = @{
            Detection = @{
                Type = "RegistryKeyValue"
                RegistryKeyRelative = "SOFTWARE\Vendor\App"
                ValueName = "DisplayVersion"
                ExpectedValue = "1.0.0"
            }
        } | ConvertTo-Json -Depth 5
        $path = Join-Path $testManifestDir "no-hive.json"
        Set-Content -LiteralPath $path -Value $manifest -Encoding UTF8

        $det = Import-DetectionManifest -Path $path
        $det.Hive | Should -Be "LocalMachine"
    }

    It "RegistryKey type also normalizes Hive" {
        $manifest = @{
            Detection = @{
                Type = "RegistryKey"
                RegistryKeyRelative = "Software\Vendor"
                Hive = "CurrentUser"
            }
        } | ConvertTo-Json -Depth 5
        $path = Join-Path $testManifestDir "hkcu-rk.json"
        Set-Content -LiteralPath $path -Value $manifest -Encoding UTF8

        $det = Import-DetectionManifest -Path $path
        $det.Hive | Should -Be "CurrentUser"
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
