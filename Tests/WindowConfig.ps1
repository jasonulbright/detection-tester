# App-specific names for Window.Tests.ps1.
#
# The window is MainWindow.xaml with a content host. The shell loads one module XAML into the host and
# fills a few texts. A view is one state of the content host: the Detection Tester with one of its five
# input panels, or the Installed Applications list. Sample rows fill the grids so that the measured
# layout has real text.
$script:AppConfig = @{
    Title         = 'Detection Tester'
    Shell         = 'start-detectiontester.ps1'
    Xaml          = 'MainWindow.xaml'
    OptionsButton = 'btnOptions'
    LogLabel      = 'lblLogOutput'
    ContentHost   = 'contentHost'
    SidebarButtons = @('btnDetection', 'btnInstalled', 'btnOptions')
    Modules       = @{
        DetectionTester = @{
            Xaml     = 'Modules\DetectionTester.xaml'
            Title    = 'Detection Tester'
            Subtitle = 'Author and test Configuration Manager detection methods locally.'
            Status   = 'Detection Tester selected.'
        }
        InstalledApps   = @{
            Xaml     = 'Modules\InstalledApps.xaml'
            Title    = 'Installed Applications'
            Subtitle = 'Browse ARP entries (HKLM and HKCU, x64 and x86).'
            Status   = 'Installed Applications selected.'
        }
    }
    Panels        = @('panelRegistryKeyValue', 'panelRegistryKey', 'panelFile', 'panelScript', 'panelCompound')
    Views         = @(
        @{ Name = 'Detection Tester: RegistryKeyValue'; Module = 'DetectionTester'; Panel = 'panelRegistryKeyValue'; TypeIndex = 0 }
        @{ Name = 'Detection Tester: RegistryKey';      Module = 'DetectionTester'; Panel = 'panelRegistryKey';      TypeIndex = 1 }
        @{ Name = 'Detection Tester: File';             Module = 'DetectionTester'; Panel = 'panelFile';             TypeIndex = 2 }
        @{ Name = 'Detection Tester: Script';           Module = 'DetectionTester'; Panel = 'panelScript';           TypeIndex = 3 }
        @{ Name = 'Detection Tester: Compound';         Module = 'DetectionTester'; Panel = 'panelCompound';         TypeIndex = 4 }
        @{ Name = 'Installed Applications';             Module = 'InstalledApps' }
    )
    SampleHistory = @(
        @{ Time = '10:15:01'; Type = 'RegistryKeyValue'; Mode = 'Positive'; Target = 'HKLM\SOFTWARE\Contoso\Application'; Expected = '1.2.3'; Found = '1.2.3'; Result = 'Pass' }
        @{ Time = '10:14:12'; Type = 'File'; Mode = 'Negative'; Target = 'C:\Program Files\Contoso\Application\application.exe'; Expected = '(file exists)'; Found = ''; Result = 'Fail' }
        @{ Time = '10:12:40'; Type = 'Compound'; Mode = 'Positive'; Target = '(And: 3 clauses)'; Expected = '(all And)'; Found = '3/3 passed'; Result = 'Pass' }
    )
    SampleCompound = @(
        "RKV:  LocalMachine\SOFTWARE\Contoso\Application [DisplayVersion] IsEquals '1.2.3'"
        'RK:   CurrentUser\Software\Contoso'
        "FILE: C:\Program Files\Contoso\application.exe (Version GreaterEquals '1.2.3.4')"
    )
    SampleApps    = @(
        @{ DisplayName = 'Contoso Application 1.2.3'; Publisher = 'Contoso Corporation'; DisplayVersion = '1.2.3.4'; Architecture = 'x86'; Scope = 'Machine'; Hive = 'LocalMachine'; RegistryKey = 'HKLM\SOFTWARE\WOW6432Node\Contoso'; RegistryKeyRelative = 'SOFTWARE\WOW6432Node\Contoso' }
        @{ DisplayName = 'Fabrikam Reporting Service with a very long product name that fills the column'; Publisher = 'Fabrikam, Inc.'; DisplayVersion = '10.0.19041.1234'; Architecture = 'x64'; Scope = 'Machine'; Hive = 'LocalMachine'; RegistryKey = 'HKLM\SOFTWARE\Fabrikam'; RegistryKeyRelative = 'SOFTWARE\Fabrikam' }
        @{ DisplayName = 'Per-user Editor'; Publisher = ''; DisplayVersion = '3.1'; Architecture = 'x64'; Scope = 'User'; Hive = 'CurrentUser'; RegistryKey = 'HKCU\Software\Editor'; RegistryKeyRelative = 'Software\Editor' }
    )
}
