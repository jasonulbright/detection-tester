# The guide steps. main.ps1 dot-sources this file in the interactive session of the lab workstation.
# $Capture holds the sections to run (Invoke-GuideCapture.ps1). The steps read the registry and the file system of the workstation
# and run detection tests on it. No step changes the workstation, except these files of the lab account: the preferences and the
# window state (reset before and after the run), the history exports in %LOCALAPPDATA%\DetectionTester\reports (removed at the end),
# and two manifest files in the capture folder.
# A section uses the state that the sections before it leave (the history rows, the page). Select every section that an image needs,
# in the order of this file.
$captureStart = Get-Date
$dataDir = Join-Path $env:LOCALAPPDATA 'DetectionTester'
function Reset-AppFiles { foreach ($f in 'prefs.json', 'windowstate.json') { Remove-Item -LiteralPath (Join-Path $dataDir $f) -Force -ErrorAction SilentlyContinue } }
Stop-App 'start-detectiontester.ps1'
Reset-AppFiles
$win = Start-App 'start-detectiontester.ps1' 'Detection Tester'
Size-Win $win 1440 900
$win = Find-Top 'Detection Tester' 10

# ---- helpers -------------------------------------------------------------------------------------------------------------------
function Show-View([string]$Button) { Invoke-El (Get-El $win $Button); Start-Sleep -Milliseconds 800 }
# Opens a list, clicks the item with this name, and leaves the list closed.
function Select-Combo([string]$Id, [string]$Item) {
    Expand-El (Get-El $win $Id); Start-Sleep -Milliseconds 500
    Click-El (Find-Role $win 'ListItem' $Item); Start-Sleep -Milliseconds 600
}
# Opens a list for an image and leaves it open until Close-Combo.
function Open-Combo([string]$Id) { Expand-El (Get-El $win $Id); Start-Sleep -Milliseconds 700 }
function Close-Combo([string]$Id) { Collapse-El (Get-El $win $Id); Start-Sleep -Milliseconds 500 }
function Set-Field([string]$Id, [string]$Text) { Set-Text (Get-El $win $Id) $Text }
function Set-Type([string]$Item) { Select-Combo 'cboType' $Item }
# Runs the test and waits for a result in the badge. A script test runs in the background.
function Invoke-Test([string]$Expect) {
    Invoke-El (Get-El $win 'btnTest'); Start-Sleep -Milliseconds 800
    $ok = Wait-Until { $t = (Get-El $win 'txtResultBadge').Current.Name; $t -ne 'Running...' } 40
    if (-not $ok) { throw 'The test gave no result.' }
    Start-Sleep -Milliseconds 700
    $got = (Get-El $win 'txtResultBadge').Current.Name
    if ($Expect -and $got -ne $Expect) { throw ('The test shows {0}, the step expects {1}.' -f $got, $Expect) }
}
function Reset-Form {
    foreach ($id in 'rkvKeyPath', 'rkvValueName', 'rkvExpected', 'rkKeyPath', 'fileFilePath', 'fileFileName', 'fileExpected', 'scriptText') { Set-Field $id '' }
}
function Get-Rows($Grid) { @($Grid.FindAll('Descendants', (New-Object System.Windows.Automation.PropertyCondition($AE::ControlTypeProperty, [System.Windows.Automation.ControlType]::DataItem)))) }
function Wait-InstalledRows {
    $null = Wait-Until { $g = Get-El $win 'dgApps' -Quiet; $g -and (Get-Rows $g).Count -gt 0 } 60
    Start-Sleep -Seconds 1
}
# The Options dialog is a window of its own with the title Options. Cancel closes it.
function Open-Options { Invoke-El (Get-El $win 'btnOptions'); Find-Top 'Options' 10 }
function Close-Options($Dlg) { Invoke-El (Get-El $Dlg 'btnCancel'); Start-Sleep -Milliseconds 600 }

$key = 'SOFTWARE\Microsoft\Windows NT\CurrentVersion'
$build = [string](Get-ItemProperty -LiteralPath 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion').CurrentBuildNumber
$sys = 'C:\Windows\System32'
# Two manifests for the import images, in the capture folder (removed with it).
$manifestOk = Join-Path $appDir 'ui\stage-manifest.json'
$manifestNone = Join-Path $appDir 'ui\stage-manifest-nodetection.json'
$clauses = @(
    [ordered]@{ Type = 'RegistryKey'; Hive = 'LocalMachine'; RegistryKeyRelative = $key; Is64Bit = $true }
    [ordered]@{ Type = 'File'; FilePath = $sys; FileName = 'notepad.exe'; PropertyType = 'Existence'; Operator = 'IsEquals'; ExpectedValue = ''; Is64Bit = $true }
)
[ordered]@{ Detection = [ordered]@{ Connector = 'And'; Clauses = $clauses } } | ConvertTo-Json -Depth 6 | Set-Content -LiteralPath $manifestOk -Encoding UTF8
[ordered]@{ Name = 'No detection here' } | ConvertTo-Json | Set-Content -LiteralPath $manifestNone -Encoding UTF8

# ---- sections ------------------------------------------------------------------------------------------------------------------
Section 'first run' {
    Show-View 'btnDetection'; Shot $win 'start-first-run'
}

Section 'detection tests' {
    Show-View 'btnDetection'
    Open-Combo 'cboType'; Shot $win 'type-list'; Close-Combo 'cboType'
    # RegistryKeyValue
    Set-Type 'RegistryKeyValue'
    Set-Field 'rkvKeyPath' $key; Set-Field 'rkvValueName' 'CurrentBuildNumber'; Set-Field 'rkvExpected' $build
    Invoke-Test 'Pass'; Shot $win 'registry-value-pass'
    Open-Combo 'rkvOperator'; Shot $win 'operator-list'; Close-Combo 'rkvOperator'
    Set-Field 'rkvExpected' '99999'; Invoke-Test 'Fail'; Shot $win 'registry-value-fail'
    Set-Field 'rkvValueName' 'NoSuchValue'; Invoke-Test 'Fail'; Shot $win 'registry-value-missing'
    Set-Field 'rkvValueName' 'CurrentBuildNumber'; Set-Field 'rkvExpected' $build
    Select-El (Get-El $win 'rkvHiveHKCU'); Select-El (Get-El $win 'rkvView32'); Invoke-Test 'Fail'; Shot $win 'registry-value-hkcu-32'
    Select-El (Get-El $win 'rkvHiveHKLM'); Select-El (Get-El $win 'rkvView64')
    # RegistryKey
    Set-Type 'RegistryKey'; Shot $win 'type-registry-key'
    Set-Field 'rkKeyPath' $key; Invoke-Test 'Pass'; Shot $win 'registry-key-pass'
    Set-Field 'rkKeyPath' 'SOFTWARE\Contoso\NoSuchKey'; Invoke-Test 'Fail'; Shot $win 'registry-key-fail'
    # File
    Set-Type 'File'; Shot $win 'type-file'
    Open-Combo 'filePropType'; Shot $win 'file-property-list'; Close-Combo 'filePropType'
    Set-Field 'fileFilePath' $sys; Set-Field 'fileFileName' 'notepad.exe'
    Invoke-Test 'Pass'; Shot $win 'file-existence-pass'
    Select-Combo 'filePropType' 'Version'; Select-Combo 'fileOperator' 'GreaterEquals'; Set-Field 'fileExpected' '10.0.0.0'
    Invoke-Test 'Pass'; Shot $win 'file-version-pass'
    Set-Field 'fileFileName' 'nosuchfile.exe'; Invoke-Test 'Fail'; Shot $win 'file-fail'
    # Script
    Set-Type 'Script'; Shot $win 'type-script'
    Set-Field 'scriptText' "Write-Output 'installed'"; Invoke-Test 'Pass'; Shot $win 'script-pass'
    Set-Field 'scriptText' '$null = 1'; Invoke-Test 'Fail'; Shot $win 'script-fail'
    # An empty field: the clause does not run.
    Set-Type 'RegistryKey'; Set-Field 'rkKeyPath' ''; Invoke-Test 'Invalid clause'; Shot $win 'test-empty'
}

Section 'history and import' {
    Show-View 'btnDetection'
    # The page needs several rows in the history before the Compound row.
    Set-Type 'RegistryKey'; Set-Field 'rkKeyPath' $key; Invoke-Test 'Pass'
    Set-Type 'Compound (import-only)'
    Invoke-El (Get-El $win 'btnImportManifest'); File-Dialog 'Import stage-manifest.json' $manifestOk
    Start-Sleep -Seconds 1; Shot $win 'import-compound'
    Invoke-Test 'Pass'; Shot $win 'import-compound-run'
    Shot $win 'history-filled'
    Invoke-El (Get-El $win 'btnImportManifest'); File-Dialog 'Import stage-manifest.json' $manifestNone
    Start-Sleep -Seconds 1; Shot $win 'import-failed'
    # Use Selected on the Compound row (the first row of the grid).
    $first = (Get-Rows (Get-El $win 'dgHistory'))[0]
    Select-El $first; Start-Sleep -Milliseconds 600
    Invoke-El (Get-El $win 'btnRepopulate'); Start-Sleep -Milliseconds 800; Shot $win 'use-selected'
    $export = Join-Path $dataDir 'reports\detection-history-capture.csv'
    Invoke-El (Get-El $win 'btnExportCsv'); File-Dialog 'Save As' $export
    Start-Sleep -Seconds 1; Shot $win 'history-after-export'
}

Section 'negative test' {
    Show-View 'btnDetection'
    # An empty history: the image shows the Negative rows only.
    Invoke-El (Get-El $win 'btnClearHistory'); Start-Sleep -Milliseconds 600
    Set-Type 'RegistryKey'; Set-Field 'rkKeyPath' $key
    Toggle-El (Get-El $win 'chkNegativeMode')
    Invoke-Test 'Fail'; Shot $win 'negative-test-fail'
    Set-Field 'rkKeyPath' 'SOFTWARE\Contoso\NoSuchKey'; Invoke-Test 'Pass'; Shot $win 'negative-test-pass'
    Toggle-El (Get-El $win 'chkNegativeMode')
}

Section 'installed applications' {
    Show-View 'btnInstalled'; Wait-InstalledRows
    Set-Text (Get-El $win 'txtFilter') 'micro'; Start-Sleep -Seconds 1; Shot $win 'installed-filter'
    Open-Combo 'cboScope'; Shot $win 'installed-scope-list'; Close-Combo 'cboScope'
    Set-Text (Get-El $win 'txtFilter') ''; Start-Sleep -Seconds 1
    Shot $win 'installed-panel'
    $rows = Get-Rows (Get-El $win 'dgApps')
    $row = @($rows | Where-Object { $_.Current.Name -like '*Configuration Manager Client*' }) | Select-Object -First 1
    if (-not $row) { $row = $rows[0] }
    Select-El $row; Start-Sleep -Seconds 1; Shot $win 'installed-details'
    Invoke-El (Get-El $win 'btnUseForDetection'); Start-Sleep -Seconds 1; Shot $win 'use-for-detection'
}

Section 'options' {
    $d = Open-Options; Shot $d 'options-about'
    Select-El (Find-Role $d 'ListItem' 'Logging'); Start-Sleep -Milliseconds 600; Shot $d 'options-logging'
    Close-Options $d
}

# The tool keeps the theme and the page that was open. Close it on the Installed Applications page and start it again.
Section 'restart' {
    Show-View 'btnInstalled'; Wait-InstalledRows
    Close-Win $win
    $null = Wait-Until { -not (Find-All 'Detection Tester').Count } 20
    Start-Sleep -Seconds 2
    Stop-App 'start-detectiontester.ps1'
    $script:win = Start-App 'start-detectiontester.ps1' 'Detection Tester'
    Size-Win $script:win 1440 900
    $script:win = Find-Top 'Detection Tester' 10
    Wait-InstalledRows
    Shot $script:win 'start-reopen'
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
Reset-AppFiles
Get-ChildItem -LiteralPath (Join-Path $dataDir 'reports') -Filter 'detection-history-*' -ErrorAction SilentlyContinue | Where-Object { $_.LastWriteTime -ge $captureStart } | Remove-Item -Force -ErrorAction SilentlyContinue
