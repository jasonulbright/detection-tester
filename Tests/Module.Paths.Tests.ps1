#Requires -Modules @{ ModuleName = 'Pester'; ModuleVersion = '5.1.0' }
# Hive and registry path helpers: every prefix form, empty and hostile input.

BeforeAll {
    . "$PSScriptRoot\TestHelpers.ps1"
    Import-AppModule
    Mock Write-Log {} -ModuleName DetectionTesterCommon
}

Describe 'ConvertTo-DTHiveName' {
    It 'maps "<Value>" to <Hive>' -ForEach @(
        @{ Value = 'HKLM';                Hive = 'LocalMachine' }
        @{ Value = 'hklm';                Hive = 'LocalMachine' }
        @{ Value = 'HKEY_LOCAL_MACHINE';  Hive = 'LocalMachine' }
        @{ Value = 'localmachine';        Hive = 'LocalMachine' }
        @{ Value = ' HKLM ';              Hive = 'LocalMachine' }
        @{ Value = 'HKCU';                Hive = 'CurrentUser' }
        @{ Value = 'hkcu';                Hive = 'CurrentUser' }
        @{ Value = 'HKEY_CURRENT_USER';   Hive = 'CurrentUser' }
        @{ Value = 'currentuser';         Hive = 'CurrentUser' }
        @{ Value = "`tHKCU`r`n";          Hive = 'CurrentUser' }
        # Missing and unknown values fall back to LocalMachine, the ConfigMgr default.
        @{ Value = '';                    Hive = 'LocalMachine' }
        @{ Value = '   ';                 Hive = 'LocalMachine' }
        @{ Value = $null;                 Hive = 'LocalMachine' }
        @{ Value = 'HKEY_USERS';          Hive = 'LocalMachine' }
        @{ Value = 'HKCUX';               Hive = 'LocalMachine' }
    ) {
        InModuleScope DetectionTesterCommon -Parameters @{ Value = $Value; Hive = $Hive } {
            ConvertTo-DTHiveName -Value $Value | Should -Be $Hive
        }
    }
}

Describe 'Get-DTHiveDisplayPrefix' {
    It 'gives <Prefix> for <Hive>' -ForEach @(
        @{ Hive = 'LocalMachine'; Prefix = 'HKLM' }
        @{ Hive = 'CurrentUser';  Prefix = 'HKCU' }
    ) {
        InModuleScope DetectionTesterCommon -Parameters @{ Hive = $Hive; Prefix = $Prefix } {
            Get-DTHiveDisplayPrefix -Hive $Hive | Should -Be $Prefix
        }
    }

    It 'rejects a hive outside the validate set' {
        InModuleScope DetectionTesterCommon {
            { Get-DTHiveDisplayPrefix -Hive 'ClassesRoot' } | Should -Throw
        }
    }
}

Describe 'ConvertTo-DTRegistryRelativePath' {
    It 'turns "<Path>" into "<Relative>"' -ForEach @(
        @{ Path = 'HKLM\SOFTWARE\Vendor';                                   Relative = 'SOFTWARE\Vendor' }
        @{ Path = 'hklm\software\vendor';                                   Relative = 'software\vendor' }
        @{ Path = 'HKEY_LOCAL_MACHINE\SOFTWARE\Vendor';                     Relative = 'SOFTWARE\Vendor' }
        @{ Path = 'HKLM:\SOFTWARE\Vendor';                                  Relative = 'SOFTWARE\Vendor' }
        @{ Path = 'HKLM:SOFTWARE\Vendor';                                   Relative = 'SOFTWARE\Vendor' }
        @{ Path = 'HKCU\Software\Vendor';                                   Relative = 'Software\Vendor' }
        @{ Path = 'HKEY_CURRENT_USER\Software\Vendor';                      Relative = 'Software\Vendor' }
        @{ Path = 'HKCU:\Software\Vendor';                                  Relative = 'Software\Vendor' }
        @{ Path = 'Registry::HKLM\SOFTWARE\Vendor';                         Relative = 'SOFTWARE\Vendor' }
        @{ Path = 'Registry::HKEY_LOCAL_MACHINE\SOFTWARE\Vendor';           Relative = 'SOFTWARE\Vendor' }
        @{ Path = 'Microsoft.PowerShell.Core\Registry::HKCU\Software\V';    Relative = 'Software\V' }
        @{ Path = '  HKLM\SOFTWARE\Vendor  ';                               Relative = 'SOFTWARE\Vendor' }
        @{ Path = '\SOFTWARE\Vendor';                                       Relative = 'SOFTWARE\Vendor' }
        @{ Path = 'SOFTWARE\Vendor';                                        Relative = 'SOFTWARE\Vendor' }
        @{ Path = 'SOFTWARE\Vendor\';                                       Relative = 'SOFTWARE\Vendor\' }
        @{ Path = 'HKLMX\Vendor';                                           Relative = 'HKLMX\Vendor' }
        @{ Path = 'HKLM_OLD\Vendor';                                        Relative = 'HKLM_OLD\Vendor' }
        @{ Path = 'HKEY_USERS\S-1-5-18\Software';                           Relative = 'HKEY_USERS\S-1-5-18\Software' }
        @{ Path = 'SOFTWARE\HKLM\Vendor';                                   Relative = 'SOFTWARE\HKLM\Vendor' }
    ) {
        InModuleScope DetectionTesterCommon -Parameters @{ Path = $Path; Relative = $Relative } {
            ConvertTo-DTRegistryRelativePath -Path $Path | Should -Be $Relative
        }
    }

    It 'gives an empty path for a path that is only a hive name: "<Path>"' -ForEach @(
        @{ Path = 'HKLM' }
        @{ Path = 'HKCU' }
        @{ Path = 'HKEY_LOCAL_MACHINE' }
        @{ Path = 'HKEY_CURRENT_USER' }
        @{ Path = 'HKLM\' }
        @{ Path = 'HKLM:' }
        @{ Path = 'HKLM:\' }
        @{ Path = 'Registry::HKLM' }
        @{ Path = 'Registry::HKEY_CURRENT_USER\' }
    ) {
        InModuleScope DetectionTesterCommon -Parameters @{ Path = $Path } {
            ConvertTo-DTRegistryRelativePath -Path $Path | Should -Be ''
        }
    }

    It 'gives an empty path for empty input: <Case>' -ForEach @(
        @{ Case = 'null'; Path = $null }
        @{ Case = 'empty'; Path = '' }
        @{ Case = 'spaces'; Path = '    ' }
    ) {
        InModuleScope DetectionTesterCommon -Parameters @{ Path = $Path } {
            ConvertTo-DTRegistryRelativePath -Path $Path | Should -Be ''
        }
    }
}

Describe 'Get-DTHiveFromPath' {
    It 'reads <Hive> from "<Path>"' -ForEach @(
        @{ Path = 'HKLM\SOFTWARE\V';                                  Hive = 'LocalMachine' }
        @{ Path = 'hklm:\software\v';                                 Hive = 'LocalMachine' }
        @{ Path = 'HKEY_LOCAL_MACHINE\SOFTWARE\V';                    Hive = 'LocalMachine' }
        @{ Path = 'Registry::HKEY_LOCAL_MACHINE\SOFTWARE\V';          Hive = 'LocalMachine' }
        @{ Path = 'HKLM';                                             Hive = 'LocalMachine' }
        @{ Path = 'HKLM:';                                            Hive = 'LocalMachine' }
        @{ Path = 'HKCU\Software\V';                                  Hive = 'CurrentUser' }
        @{ Path = 'HKCU:\Software\V';                                 Hive = 'CurrentUser' }
        @{ Path = 'HKEY_CURRENT_USER\Software\V';                     Hive = 'CurrentUser' }
        @{ Path = 'Microsoft.PowerShell.Core\Registry::HKCU\Software'; Hive = 'CurrentUser' }
        @{ Path = 'HKCU';                                             Hive = 'CurrentUser' }
        @{ Path = ' HKCU\Software ';                                  Hive = 'CurrentUser' }
        @{ Path = 'SOFTWARE\Vendor';                                  Hive = '' }
        @{ Path = 'HKLMX\Vendor';                                     Hive = '' }
        @{ Path = 'SOFTWARE\HKLM\Vendor';                             Hive = '' }
        @{ Path = 'HKEY_USERS\S-1\x';                                 Hive = '' }
        @{ Path = '';                                                 Hive = '' }
        @{ Path = $null;                                              Hive = '' }
        @{ Path = '   ';                                              Hive = '' }
    ) {
        InModuleScope DetectionTesterCommon -Parameters @{ Path = $Path; Hive = $Hive } {
            Get-DTHiveFromPath -Path $Path | Should -Be $Hive
        }
    }

    It 'reads the same hive that ConvertTo-DTRegistryRelativePath strips' -ForEach @(
        @{ Path = 'HKLM:\SOFTWARE\V' }, @{ Path = 'HKCU\Software\V' }, @{ Path = 'Registry::HKEY_CURRENT_USER\Software\V' }
    ) {
        InModuleScope DetectionTesterCommon -Parameters @{ Path = $Path } {
            (Get-DTHiveFromPath -Path $Path) | Should -Not -BeNullOrEmpty
            (ConvertTo-DTRegistryRelativePath -Path $Path) | Should -Not -Match '^(HKLM|HKCU|HKEY_|Registry::)'
        }
    }
}
