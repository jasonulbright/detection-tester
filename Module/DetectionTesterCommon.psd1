@{
    RootModule        = 'DetectionTesterCommon.psm1'
    ModuleVersion     = '1.0.0'
    GUID              = 'a1b2c3d4-5678-9abc-def0-112233445566'
    Author            = 'Jason Ulbright'
    Description       = 'Shared module for Detection Method Testing Tool: detection tests, ARP enumeration, manifest import, export.'
    PowerShellVersion = '5.1'

    FunctionsToExport = @(
        # Logging
        'Initialize-Logging'
        'Write-Log'

        # Detection tests
        'Test-RegistryKeyValueDetection'
        'Test-RegistryKeyDetection'
        'Test-FileDetection'
        'Test-ScriptDetection'
        'Test-CompoundDetection'

        # ARP enumeration
        'Get-InstalledApplications'

        # Manifest import
        'Import-DetectionManifest'

        # Export
        'Export-DetectionResultsCsv'
        'Export-DetectionResultsHtml'
    )

    CmdletsToExport   = @()
    VariablesToExport  = @()
    AliasesToExport    = @()
}
