#Requires -Modules Pester
# Manifest import.

BeforeAll {
    . "$PSScriptRoot\TestHelpers.ps1"
    Import-AppModule
    Mock Write-Log {} -ModuleName DetectionTesterCommon
}
Describe 'Codex review fixes' {
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

Describe 'Import-DetectionManifest field normalization' {
    BeforeAll {
        $script:Dir = Join-Path $TestDrive 'norm'
        New-Item -ItemType Directory -Path $script:Dir -Force | Out-Null
        function script:Write-Manifest {
            param([string]$Name, [string]$Json, [switch]$Bom)
            $p = Join-Path $script:Dir $Name
            [System.IO.File]::WriteAllText($p, $Json, (New-Object System.Text.UTF8Encoding([bool]$Bom)))
            return $p
        }
    }

    It 'defaults the fields of a registry value block that has only a type' {
        $d = Import-DetectionManifest -Path (Write-Manifest 'rkv-min.json' '{"Detection":{"Type":"RegistryKeyValue","RegistryKeyRelative":"SOFTWARE\\V"}}')
        $d.Type | Should -Be 'RegistryKeyValue'
        $d.ValueName | Should -Be 'DisplayVersion'
        $d.ExpectedValue | Should -Be ''
        $d.Operator | Should -Be 'IsEquals'
        $d.PropertyType | Should -Be 'String'
        $d.Is64Bit | Should -BeTrue
        $d.Hive | Should -Be 'LocalMachine'
    }

    It 'treats a block without a type as a registry value block' {
        $d = Import-DetectionManifest -Path (Write-Manifest 'notype.json' '{"Detection":{"KeyPath":"HKCU\\Software\\V","ValueName":"x","ExpectedValue":"1"}}')
        $d.Type | Should -Be 'RegistryKeyValue'
        $d.Hive | Should -Be 'CurrentUser'
        $d.RegistryKeyRelative | Should -Be 'Software\V'
    }

    It 'reads an expected value of 0 as "0"' -ForEach @(
        @{ Json = '{"Detection":{"Type":"RegistryKeyValue","RegistryKeyRelative":"S","ValueName":"Dw","ExpectedValue":0}}' }
        @{ Json = '{"Detection":{"Type":"RegistryKeyValue","RegistryKeyRelative":"S","ValueName":"Dw","ExpectedValue":"0"}}' }
    ) {
        (Import-DetectionManifest -Path (Write-Manifest 'zero.json' $Json)).ExpectedValue | Should -Be '0'
    }

    It 'reads an expected value of 0 in a file block as "0"' {
        $d = Import-DetectionManifest -Path (Write-Manifest 'filezero.json' '{"Detection":{"Type":"File","FilePath":"C:\\x","FileName":"a.exe","PropertyType":"Version","ExpectedValue":0}}')
        $d.ExpectedValue | Should -Be '0'
    }

    It 'prefers ExpectedValue over DisplayVersion, and uses DisplayVersion when ExpectedValue is missing' {
        (Import-DetectionManifest -Path (Write-Manifest 'both.json' '{"Detection":{"Type":"RegistryKeyValue","RegistryKeyRelative":"S","ExpectedValue":"1","DisplayVersion":"2"}}')).ExpectedValue | Should -Be '1'
        (Import-DetectionManifest -Path (Write-Manifest 'dv.json' '{"Detection":{"Type":"RegistryKeyValue","RegistryKeyRelative":"S","DisplayVersion":"2"}}')).ExpectedValue | Should -Be '2'
    }

    It 'prefers RegistryKeyRelative over KeyPath' {
        $d = Import-DetectionManifest -Path (Write-Manifest 'both-path.json' '{"Detection":{"Type":"RegistryKey","RegistryKeyRelative":"A\\B","KeyPath":"C\\D"}}')
        $d.RegistryKeyRelative | Should -Be 'A\B'
    }

    It 'lets an explicit Hive win over the prefix of the path' {
        $d = Import-DetectionManifest -Path (Write-Manifest 'hive-wins.json' '{"Detection":{"Type":"RegistryKey","KeyPath":"HKCU\\Software\\V","Hive":"HKLM"}}')
        $d.Hive | Should -Be 'LocalMachine'
        $d.RegistryKeyRelative | Should -Be 'Software\V'
    }

    It 'reads Is64Bit <Raw> as <Expected>' -ForEach @(
        @{ Raw = 'true';    Expected = $true }
        @{ Raw = 'false';   Expected = $false }
        @{ Raw = '"false"'; Expected = $false }
        @{ Raw = '"true"';  Expected = $true }
        @{ Raw = '0';       Expected = $false }
        @{ Raw = '1';       Expected = $true }
        @{ Raw = 'null';    Expected = $true }
    ) {
        $d = Import-DetectionManifest -Path (Write-Manifest 'is64.json' ('{"Detection":{"Type":"RegistryKey","RegistryKeyRelative":"S","Is64Bit":' + $Raw + '}}'))
        $d.Is64Bit | Should -Be $Expected
    }

    It 'defaults the fields of a file block' {
        $d = Import-DetectionManifest -Path (Write-Manifest 'file-min.json' '{"Detection":{"Type":"File","FilePath":"C:\\x","FileName":"a.exe"}}')
        $d.PropertyType | Should -Be 'Existence'
        $d.Operator | Should -Be 'GreaterEquals'
        $d.ExpectedValue | Should -Be ''
        $d.Is64Bit | Should -BeTrue
    }

    It 'defaults the fields of a script block' {
        $d = Import-DetectionManifest -Path (Write-Manifest 'script-min.json' '{"Detection":{"Type":"Script","ScriptText":"\"x\""}}')
        $d.ScriptText | Should -Be '"x"'
        $d.ScriptLanguage | Should -Be 'PowerShell'
        (Import-DetectionManifest -Path (Write-Manifest 'script-empty.json' '{"Detection":{"Type":"Script"}}')).ScriptText | Should -Be ''
    }

    It 'returns only the type for a block type it does not know' {
        $d = Import-DetectionManifest -Path (Write-Manifest 'unknown-type.json' '{"Detection":{"Type":"WmiQuery","Query":"select 1"}}')
        $d.Type | Should -Be 'WmiQuery'
        $d.Keys.Count | Should -Be 1
    }

    It 'reads a compound block and normalizes each clause' {
        $json = '{"Detection":{"Connector":"Or","Clauses":[{"Type":"RegistryKey","KeyPath":"HKCU:\\Software\\A"},{"Type":"File","FilePath":"C:\\x","FileName":"a.exe"},{"KeyPath":"HKLM\\SOFTWARE\\B","ValueName":"v","ExpectedValue":0}]}}'
        $d = Import-DetectionManifest -Path (Write-Manifest 'compound.json' $json)
        $d.Type | Should -Be 'Compound'
        $d.Connector | Should -Be 'Or'
        $d.Clauses.Count | Should -Be 3
        $d.Clauses[0].Hive | Should -Be 'CurrentUser'
        $d.Clauses[0].RegistryKeyRelative | Should -Be 'Software\A'
        $d.Clauses[1].PropertyType | Should -Be 'Existence'
        $d.Clauses[2].Type | Should -Be 'RegistryKeyValue'
        $d.Clauses[2].ExpectedValue | Should -Be '0'
    }

    It 'reads a compound block with one clause as a list of one' {
        $d = Import-DetectionManifest -Path (Write-Manifest 'compound1.json' '{"Detection":{"Connector":"And","Clauses":[{"Type":"RegistryKey","RegistryKeyRelative":"S"}]}}')
        $d.Type | Should -Be 'Compound'
        @($d.Clauses).Count | Should -Be 1
    }

    It 'reads a block that has a connector but no clauses as a single block' {
        $d = Import-DetectionManifest -Path (Write-Manifest 'connector-only.json' '{"Detection":{"Connector":"And","Type":"RegistryKey","RegistryKeyRelative":"S"}}')
        $d.Type | Should -Be 'RegistryKey'
    }

    It 'reads a manifest with non-ASCII text, with and without a byte order mark: <Case>' -ForEach @(
        @{ Case = 'no byte order mark'; Bom = $false }
        @{ Case = 'byte order mark';    Bom = $true }
    ) {
        $text = "C:\Program Files\Caf$([char]0xE9) $([char]0x20AC)"
        $json = '{"Detection":{"Type":"File","FilePath":"' + $text.Replace('\', '\\') + '","FileName":"a.exe"}}'
        $path = Write-Manifest 'nonascii.json' $json -Bom:$Bom
        (Import-DetectionManifest -Path $path).FilePath | Should -Be $text
    }

    It 'reads a manifest from a path with brackets' {
        $dir = Join-Path $script:Dir 'a[1]'
        New-Item -ItemType Directory -Path $dir -Force | Out-Null
        $p = Join-Path $dir 'm[1].json'
        [System.IO.File]::WriteAllText($p, '{"Detection":{"Type":"RegistryKey","RegistryKeyRelative":"S"}}')
        (Import-DetectionManifest -Path $p).RegistryKeyRelative | Should -Be 'S'
    }

    It 'throws "Failed to parse manifest JSON" for <Case>' -ForEach @(
        @{ Case = 'text that is not JSON'; Json = 'not json' }
        @{ Case = 'JSON that stops early'; Json = '{"Detection":{"Type":' }
    ) {
        { Import-DetectionManifest -Path (Write-Manifest 'bad.json' $Json) } | Should -Throw '*Failed to parse manifest JSON*'
    }

    It 'throws for an empty file' {
        { Import-DetectionManifest -Path (Write-Manifest 'empty.json' '') } | Should -Throw
    }

    It 'throws "No Detection block" for <Case>' -ForEach @(
        @{ Case = 'an object without Detection'; Json = '{"AppName":"x"}' }
        @{ Case = 'a null Detection';            Json = '{"Detection":null}' }
    ) {
        { Import-DetectionManifest -Path (Write-Manifest 'nodet2.json' $Json) } | Should -Throw "*No 'Detection' block*"
    }

    It 'throws "not found" for a folder path that does not exist and for a missing file name' {
        { Import-DetectionManifest -Path (Join-Path $script:Dir 'nope\m.json') } | Should -Throw '*Manifest file not found*'
        { Import-DetectionManifest -Path (Join-Path $script:Dir 'nope.json') } | Should -Throw '*Manifest file not found*'
    }

    It 'reports a read error as a read error, not as a missing Detection block' {
        { Import-DetectionManifest -Path $script:Dir } | Should -Throw '*Failed to read manifest*'
    }

    It 'reports a locked file as a read error' {
        $p = Write-Manifest 'locked.json' '{"Detection":{"Type":"RegistryKey","RegistryKeyRelative":"S"}}'
        $stream = [System.IO.File]::Open($p, [System.IO.FileMode]::Open, [System.IO.FileAccess]::ReadWrite, [System.IO.FileShare]::None)
        try { { Import-DetectionManifest -Path $p } | Should -Throw '*Failed to read manifest*' }
        finally { $stream.Dispose() }
    }

    It 'keeps a script block text that looks like an injection as plain text' {
        $text = '"; Remove-Item C:\ -Recurse #'
        $json = (@{ Detection = @{ Type = 'Script'; ScriptText = $text } } | ConvertTo-Json -Depth 4)
        (Import-DetectionManifest -Path (Write-Manifest 'inject.json' $json)).ScriptText | Should -Be $text
    }
}

Describe 'Resolve-DetectionFields' {
    It 'normalizes an object that came from ConvertFrom-Json and a hashtable alike' {
        InModuleScope DetectionTesterCommon {
            $fromHash = Resolve-DetectionFields -Det @{ Type = 'RegistryKey'; KeyPath = 'HKCU\Software\V' }
            $fromObject = Resolve-DetectionFields -Det ([pscustomobject]@{ Type = 'RegistryKey'; KeyPath = 'HKCU\Software\V' })
            $fromHash.Hive | Should -Be $fromObject.Hive
            $fromHash.RegistryKeyRelative | Should -Be $fromObject.RegistryKeyRelative
        }
    }

    It 'gives an empty key path for a path that is only a hive name' {
        InModuleScope DetectionTesterCommon {
            $r = Resolve-DetectionFields -Det ([pscustomobject]@{ Type = 'RegistryKey'; KeyPath = 'HKLM' })
            $r.Hive | Should -Be 'LocalMachine'
            $r.RegistryKeyRelative | Should -Be ''
        }
    }

    It 'keeps a number zero and a false boolean as values, not as missing' {
        InModuleScope DetectionTesterCommon {
            $r = Resolve-DetectionFields -Det ([pscustomobject]@{ Type = 'RegistryKeyValue'; RegistryKeyRelative = 'S'; ExpectedValue = 0; Is64Bit = $false })
            $r.ExpectedValue | Should -Be '0'
            $r.Is64Bit | Should -BeFalse
        }
    }
}
