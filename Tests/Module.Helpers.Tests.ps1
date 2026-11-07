#Requires -Modules Pester
# Module loading, value comparison and path helpers.

BeforeAll {
    . "$PSScriptRoot\TestHelpers.ps1"
    Import-AppModule
    Mock Write-Log {} -ModuleName DetectionTesterCommon
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

    It "resolves logging through the vendored SuiteCommon module" {
        # Initialize-Logging / Write-Log moved to Lib\SuiteCommon, imported
        # globally by the root module.
        (Get-Command -Module SuiteCommon).Name | Should -Contain 'Initialize-Logging'
        (Get-Command -Module SuiteCommon).Name | Should -Contain 'Write-Log'
    }
}

Describe 'Codex review fixes' {
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
}
