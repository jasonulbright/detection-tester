#Requires -Modules @{ ModuleName = 'Pester'; ModuleVersion = '5.1.0' }
# Installed application enumeration.

BeforeAll {
    . "$PSScriptRoot\TestHelpers.ps1"
    Import-AppModule
    Mock Write-Log {} -ModuleName DetectionTesterCommon
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

Describe 'Get-InstalledApplications on throwaway ARP entries' {
    BeforeAll {
        $uninstall = 'Software\Microsoft\Windows\CurrentVersion\Uninstall'
        $script:Keys = @{}
        $script:Keys.Full = New-TestRegistryKey -Parent $uninstall -Value @{
            DisplayName = 'ZZTEST-DT Full'; Publisher = 'ZZ Publisher'; DisplayVersion = '1.2.3'
            UninstallString = 'C:\zz\unins.exe /S'; QuietUninstallString = 'C:\zz\unins.exe /VERYSILENT'
            InstallLocation = 'C:\Program Files\ZZ'; InstallDate = '20261106'
        }
        $script:Keys.Minimal = New-TestRegistryKey -Parent $uninstall -Value @{ DisplayName = 'ZZTEST-DT Minimal' }
        $script:Keys.NoName = New-TestRegistryKey -Parent $uninstall -Value @{ Publisher = 'ZZTEST-DT NoName Publisher' }
        $script:Keys.BlankName = New-TestRegistryKey -Parent $uninstall -Value @{ DisplayName = '   '; Publisher = 'ZZTEST-DT Blank Publisher' }
        $script:Keys.Number = New-TestRegistryKey -Parent $uninstall -Value @{ DisplayName = 'ZZTEST-DT Number'; DisplayVersion = @{ Data = 5; Kind = 'DWord' } }
        $script:Keys.Unicode = New-TestRegistryKey -Parent $uninstall -Value @{ DisplayName = "ZZTEST-DT Caf$([char]0xE9) $([char]0x20AC)" }
        $script:Apps = @(Get-InstalledApplications)
        function script:Find-App([string]$Name) { $script:Apps | Where-Object { $_.DisplayName -eq $Name } }
    }
    AfterAll { foreach ($k in $script:Keys.Values) { Remove-TestRegistryKey -Relative $k } }

    It 'lists an entry with all its fields' {
        $a = @(Find-App 'ZZTEST-DT Full')
        $a.Count | Should -Be 1
        $a[0].Publisher | Should -Be 'ZZ Publisher'
        $a[0].DisplayVersion | Should -Be '1.2.3'
        $a[0].UninstallString | Should -Be 'C:\zz\unins.exe /S'
        $a[0].QuietUninstallString | Should -Be 'C:\zz\unins.exe /VERYSILENT'
        $a[0].InstallLocation | Should -Be 'C:\Program Files\ZZ'
        $a[0].InstallDate | Should -Be '20261106'
        $a[0].Architecture | Should -Be 'x64'
        $a[0].Scope | Should -Be 'User'
        $a[0].Hive | Should -Be 'CurrentUser'
        $a[0].RegistryKeyRelative | Should -Be ('SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\' + (Split-Path -Path $script:Keys.Full -Leaf))
        $a[0].RegistryKey | Should -Be ('HKCU\' + $a[0].RegistryKeyRelative)
    }

    It 'gives empty text for the fields an entry does not have' {
        $a = Find-App 'ZZTEST-DT Minimal'
        foreach ($f in 'Publisher', 'DisplayVersion', 'UninstallString', 'QuietUninstallString', 'InstallLocation', 'InstallDate') {
            $a.$f | Should -BeOfType [string]
            $a.$f | Should -Be ''
        }
    }

    It 'leaves out an entry without a display name and one with a blank display name' {
        @($script:Apps | Where-Object { $_.Publisher -like 'ZZTEST-DT*' }) | Should -BeNullOrEmpty
    }

    It 'shows a numeric DisplayVersion as text' {
        $a = Find-App 'ZZTEST-DT Number'
        $a.DisplayVersion | Should -BeOfType [string]
        $a.DisplayVersion | Should -Be '5'
    }

    It 'keeps non-ASCII letters in the display name' {
        @(Find-App "ZZTEST-DT Caf$([char]0xE9) $([char]0x20AC)").Count | Should -Be 1
    }

    It 'sorts the list by display name' {
        $names = @($script:Apps.DisplayName)
        $sorted = @($script:Apps | Sort-Object DisplayName | ForEach-Object DisplayName)
        ($names -join "`n") | Should -BeExactly ($sorted -join "`n")
    }

    It 'pairs scope, hive, architecture and key prefix consistently' {
        foreach ($a in $script:Apps) {
            if ($a.Scope -eq 'Machine') { $a.Hive | Should -Be 'LocalMachine'; $a.RegistryKey | Should -BeLike 'HKLM\*'; $a.Architecture | Should -BeIn 'x64', 'x86' }
            else { $a.Scope | Should -Be 'User'; $a.Hive | Should -Be 'CurrentUser'; $a.RegistryKey | Should -BeLike 'HKCU\*'; $a.Architecture | Should -Be 'x64' }
            if ($a.Architecture -eq 'x86' -and $a.Scope -eq 'Machine') { $a.RegistryKeyRelative | Should -BeLike '*WOW6432Node*' }
            $a.DisplayName | Should -Not -BeNullOrEmpty
        }
    }

    It 'gives a key that the key test finds, for each user entry and the first 40 machine entries' {
        $sample = @($script:Apps | Where-Object Scope -eq 'User') + @($script:Apps | Where-Object Scope -eq 'Machine' | Select-Object -First 40)
        $sample.Count | Should -BeGreaterThan 0
        $missing = foreach ($a in $sample) {
            $r = Test-RegistryKeyDetection -RegistryKeyRelative $a.RegistryKeyRelative -Hive $a.Hive -Is64Bit ($a.Architecture -ne 'x86')
            if (-not $r.Detected) { $a.RegistryKey }
        }
        @($missing) | Should -BeNullOrEmpty
    }

    It 'lists the throwaway entries once each' {
        @($script:Apps | Where-Object { $_.DisplayName -like 'ZZTEST-DT*' }).Count | Should -Be 4
    }
}
