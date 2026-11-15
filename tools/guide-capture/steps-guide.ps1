# The guide steps. main.ps1 dot-sources this file in the interactive session of the lab workstation.
# $Capture holds the sections to run (Invoke-GuideCapture.ps1). The steps only read: no step registers a detection method.
# The app keeps its preferences in %LOCALAPPDATA%\DetectionTester. The first lines reset them, so that the run starts in the
# first-run state (dark theme, Detection Tester page, no history), and the last lines remove them again.
Stop-App 'start-detectiontester.ps1'
foreach ($f in 'prefs.json', 'windowstate.json') { Remove-Item -LiteralPath (Join-Path $env:LOCALAPPDATA "DetectionTester\$f") -Force -ErrorAction SilentlyContinue }
$win = Start-App 'start-detectiontester.ps1' 'Detection Tester'
Size-Win $win 1440 900
$win = Find-Top 'Detection Tester' 10
function Show-View([string]$Button) { Invoke-El (Get-El $win $Button); Start-Sleep -Milliseconds 800 }
# The Options dialog is a window of its own with the title Options. Cancel closes it.
function Open-Options { Invoke-El (Get-El $win 'btnOptions'); Find-Top 'Options' 10 }
function Close-Options($Dlg) { Invoke-El (Get-El $Dlg 'btnCancel'); Start-Sleep -Milliseconds 600 }
# The installed applications list reads the registry: wait until the grid shows a row.
function Wait-InstalledRows {
    $null = Wait-Until { $g = Get-El $win 'dgApps' -Quiet; $g -and @($g.FindAll('Descendants', (New-Object System.Windows.Automation.PropertyCondition($AE::ControlTypeProperty, [System.Windows.Automation.ControlType]::DataItem)))).Count -gt 0 } 60
    Start-Sleep -Seconds 1
}

Section 'dark theme' {
    Show-View 'btnDetection'; Shot $win 'start-first-run'
    Show-View 'btnInstalled'; Wait-InstalledRows; Shot $win 'installed-panel'
    $d = Open-Options; Shot $d 'options-about'
    Select-El (Find-Role $d 'ListItem' 'Logging'); Start-Sleep -Milliseconds 600; Shot $d 'options-logging'
    Close-Options $d
}
Section 'light theme' {
    Toggle-El (Get-El $win 'toggleTheme'); Start-Sleep -Seconds 1
    Show-View 'btnDetection'; Shot $win 'light-detection'
    Show-View 'btnInstalled'; Wait-InstalledRows; Shot $win 'light-installed'
    $d = Open-Options; Shot $d 'light-options'
    Close-Options $d
    # The theme goes back to dark, so that the preference of the lab account is the default again.
    Toggle-El (Get-El $win 'toggleTheme'); Start-Sleep -Seconds 1
}
Stop-App 'start-detectiontester.ps1'
foreach ($f in 'prefs.json', 'windowstate.json') { Remove-Item -LiteralPath (Join-Path $env:LOCALAPPDATA "DetectionTester\$f") -Force -ErrorAction SilentlyContinue }
