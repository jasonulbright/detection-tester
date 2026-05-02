<#
.SYNOPSIS
    Detection Method Tester - WPF shell loader.

.DESCRIPTION
    Loads the MahApps WPF shell and dispatches to the Detection Tester and
    Installed Applications modules. Detection logic lives in
    Module\DetectionTesterCommon.psm1 (GUI-independent).

    Tests MECM application detection methods (RegistryKeyValue, RegistryKey,
    File, Script, Compound) against the local machine without deploying
    through MECM. Also browses ARP entries (HKLM and HKCU) for quick
    clause authoring.

.EXAMPLE
    .\start-detectiontester.ps1

.NOTES
    Requirements:
      - PowerShell 5.1
      - .NET Framework 4.7.2+
      - Vendored MahApps DLLs in .\Lib\

    No admin rights and no MECM connection required.

    ScriptName : start-detectiontester.ps1
    Purpose    : Test MECM detection methods locally (WPF shell)
    Version    : 1.0.0
    Updated    : 2026-05-02
#>

param()

$ErrorActionPreference = 'Stop'

# ---------------------------------------------------------------------------
# Vendored MahApps + WPF
# ---------------------------------------------------------------------------

$libDir = Join-Path $PSScriptRoot 'Lib'
Add-Type -AssemblyName PresentationFramework
Add-Type -AssemblyName PresentationCore
Add-Type -AssemblyName WindowsBase
Add-Type -AssemblyName System.Xaml
Add-Type -Path (Join-Path $libDir 'ControlzEx.dll')
Add-Type -Path (Join-Path $libDir 'MahApps.Metro.dll')
Add-Type -Path (Join-Path $libDir 'Microsoft.Xaml.Behaviors.dll')

# ---------------------------------------------------------------------------
# Module (detection logic, GUI-independent)
# ---------------------------------------------------------------------------

$moduleRoot = Join-Path $PSScriptRoot 'Module'
Import-Module (Join-Path $moduleRoot 'DetectionTesterCommon.psd1') -Force -DisableNameChecking

# ---------------------------------------------------------------------------
# Per-app data root (%LOCALAPPDATA%\DetectionTester) - logs, prefs, window state
# ---------------------------------------------------------------------------

$appDataDir = Join-Path $env:LOCALAPPDATA 'DetectionTester'
if (-not (Test-Path -LiteralPath $appDataDir)) {
    New-Item -ItemType Directory -Path $appDataDir -Force | Out-Null
}
$logDir = Join-Path $appDataDir 'logs'
if (-not (Test-Path -LiteralPath $logDir)) {
    New-Item -ItemType Directory -Path $logDir -Force | Out-Null
}
$logPath   = Join-Path $logDir ('DetectionTester-{0}.log' -f (Get-Date -Format 'yyyyMMdd-HHmmss'))
$prefsPath = Join-Path $appDataDir 'prefs.json'
$wsPath    = Join-Path $appDataDir 'windowstate.json'
$reportsDir = Join-Path $appDataDir 'reports'
if (-not (Test-Path -LiteralPath $reportsDir)) {
    New-Item -ItemType Directory -Path $reportsDir -Force | Out-Null
}

Initialize-Logging -LogPath $logPath

# ---------------------------------------------------------------------------
# Init-only functions (called from main body, never from closures)
# ---------------------------------------------------------------------------

function Get-DTPrefs {
    $defaults = @{ DarkMode = $true; ActiveModule = 'DetectionTester' }
    if (Test-Path -LiteralPath $prefsPath) {
        try {
            $loaded = Get-Content -LiteralPath $prefsPath -Raw | ConvertFrom-Json
            if ($null -ne $loaded.DarkMode)     { $defaults.DarkMode     = [bool]$loaded.DarkMode }
            if ($null -ne $loaded.ActiveModule) { $defaults.ActiveModule = [string]$loaded.ActiveModule }
        } catch {
            Write-Log "Get-DTPrefs: prefs.json malformed; using defaults ($($_.Exception.Message))" -Level WARN -Quiet
        }
    }
    return $defaults
}

function Restore-DTWindowState {
    if (-not (Test-Path -LiteralPath $wsPath)) { return }
    try {
        $s = Get-Content -LiteralPath $wsPath -Raw | ConvertFrom-Json
        $w = if ($null -ne $s.Width  -and [double]$s.Width  -gt 600) { [double]$s.Width  } else { $window.Width }
        $h = if ($null -ne $s.Height -and [double]$s.Height -gt 400) { [double]$s.Height } else { $window.Height }

        # Bounds check against the virtual screen (multi-monitor desktop).
        # If the saved Left/Top would put the window mostly off-screen
        # (e.g. monitor was unplugged, resolution changed), recenter on
        # the primary screen instead.
        $vsLeft   = [System.Windows.SystemParameters]::VirtualScreenLeft
        $vsTop    = [System.Windows.SystemParameters]::VirtualScreenTop
        $vsWidth  = [System.Windows.SystemParameters]::VirtualScreenWidth
        $vsHeight = [System.Windows.SystemParameters]::VirtualScreenHeight

        $left = if ($null -ne $s.Left) { [double]$s.Left } else { $window.Left }
        $top  = if ($null -ne $s.Top)  { [double]$s.Top  } else { $window.Top }

        # The window must overlap the virtual screen by at least 100x40 to be
        # considered visible. Otherwise center on primary.
        $visibleOverlap = ($left + $w -gt $vsLeft + 100) -and
                          ($left -lt $vsLeft + $vsWidth - 100) -and
                          ($top + $h -gt $vsTop + 40) -and
                          ($top -lt $vsTop + $vsHeight - 40)

        if (-not $visibleOverlap) {
            Write-Log "Restore-DTWindowState: saved Left=$left Top=$top off-screen; recentering" -Level WARN -Quiet
            $window.WindowStartupLocation = [System.Windows.WindowStartupLocation]::CenterScreen
            $window.Width  = $w
            $window.Height = $h
        } else {
            $window.WindowStartupLocation = [System.Windows.WindowStartupLocation]::Manual
            $window.Width  = $w
            $window.Height = $h
            $window.Left   = $left
            $window.Top    = $top
        }
        if ($s.Maximized) { $window.WindowState = [System.Windows.WindowState]::Maximized }
    } catch { Write-Log "Restore-DTWindowState: $($_.Exception.Message)" -Quiet }
}

function Save-DTPrefs {
    param([hashtable]$Prefs)
    try {
        $Prefs | ConvertTo-Json | Set-Content -LiteralPath $prefsPath -Encoding UTF8
    } catch { Write-Log "Save-DTPrefs: $($_.Exception.Message)" -Quiet }
}

function Save-DTWindowState {
    try {
        $isMax  = ($window.WindowState -eq [System.Windows.WindowState]::Maximized)
        $bounds = if ($isMax) {
            $window.RestoreBounds
        } else {
            New-Object System.Windows.Rect($window.Left, $window.Top, $window.ActualWidth, $window.ActualHeight)
        }
        $state = @{
            Left      = $bounds.Left
            Top       = $bounds.Top
            Width     = $bounds.Width
            Height    = $bounds.Height
            Maximized = $isMax
        }
        $state | ConvertTo-Json | Set-Content -LiteralPath $wsPath -Encoding UTF8
    } catch { Write-Log "Save-DTWindowState: $($_.Exception.Message)" -Quiet }
}

# Build a ComboBox->string accessor that works whether the ComboBox holds
# ComboBoxItems (XAML literals) or plain strings (programmatic).
function Get-DTComboValue {
    param([Parameter(Mandatory)]$Combo)
    $sel = $Combo.SelectedItem
    if ($null -eq $sel) { return '' }
    if ($sel -is [System.Windows.Controls.ComboBoxItem]) { return [string]$sel.Content }
    return [string]$sel
}

function Import-DTXaml {
    param([Parameter(Mandatory)][string]$Path)
    [xml]$xamlDoc = Get-Content -LiteralPath $Path -Raw
    $reader = New-Object System.Xml.XmlNodeReader $xamlDoc
    return [Windows.Markup.XamlReader]::Load($reader)
}

# =============================================================================
# Title-bar drag fallback. PS51-WPF-033.
# Some VS Code PowerShell launch contexts can leave MahApps' custom title
# thumb unable to initiate native window move. Install a WM_NCHITTEST hook
# returning HTCAPTION for the title band, plus a managed DragMove fallback
# for hosts where HwndSource cannot be hooked. Wire on every MetroWindow
# (main window and every modal popup).
# =============================================================================
$script:TitleBarHitTestWindows = @{}
$script:TitleBarHitTestHooks   = @{}

function Get-TitleBarDragHeight {
    param([MahApps.Metro.Controls.MetroWindow]$Window)
    try {
        $h = [double]$Window.TitleBarHeight
        if ($h -gt 0 -and -not [double]::IsNaN($h)) { return $h }
    } catch { $null = $_ }
    return 30.0
}

function Get-InputAncestors {
    param([System.Windows.DependencyObject]$Start)
    $cur = $Start
    while ($cur) {
        $cur
        $parent = $null
        if ($cur -is [System.Windows.Media.Visual] -or $cur -is [System.Windows.Media.Media3D.Visual3D]) {
            try { $parent = [System.Windows.Media.VisualTreeHelper]::GetParent($cur) } catch { $parent = $null }
        }
        if (-not $parent -and $cur -is [System.Windows.FrameworkElement]) { $parent = $cur.Parent }
        if (-not $parent -and $cur -is [System.Windows.FrameworkContentElement]) { $parent = $cur.Parent }
        if (-not $parent -and $cur -is [System.Windows.ContentElement]) {
            try { $parent = [System.Windows.ContentOperations]::GetParent($cur) } catch { $parent = $null }
        }
        $cur = $parent
    }
}

function Test-IsWindowCommandPoint {
    param([MahApps.Metro.Controls.MetroWindow]$Window, [System.Windows.Point]$Point)
    try {
        [void]$Window.ApplyTemplate()
        $commands = $Window.Template.FindName('PART_WindowButtonCommands', $Window)
        if ($commands -and $commands.IsVisible -and $commands.ActualWidth -gt 0 -and $commands.ActualHeight -gt 0) {
            $origin = $commands.TransformToAncestor($Window).Transform([System.Windows.Point]::new(0, 0))
            if ($Point.X -ge $origin.X -and $Point.X -le ($origin.X + $commands.ActualWidth) -and
                $Point.Y -ge $origin.Y -and $Point.Y -le ($origin.Y + $commands.ActualHeight)) {
                return $true
            }
        }
    } catch { $null = $_ }
    return ($Window.ActualWidth -gt 150 -and $Point.X -ge ($Window.ActualWidth - 150))
}

function Add-NativeTitleBarHitTestHook {
    param([MahApps.Metro.Controls.MetroWindow]$Window)
    try {
        $helper = [System.Windows.Interop.WindowInteropHelper]::new($Window)
        $source = [System.Windows.Interop.HwndSource]::FromHwnd($helper.Handle)
        if (-not $source) { return }
        $key = $helper.Handle.ToInt64().ToString()
        if ($script:TitleBarHitTestHooks.ContainsKey($key)) { return }
        $script:TitleBarHitTestWindows[$key] = $Window
        $hook = [System.Windows.Interop.HwndSourceHook]{
            param([IntPtr]$hwnd, [int]$msg, [IntPtr]$wParam, [IntPtr]$lParam, [ref]$handled)
            $WM_NCHITTEST = 0x0084; $HTCAPTION = 2
            if ($msg -ne $WM_NCHITTEST) { return [IntPtr]::Zero }
            try {
                $target = $script:TitleBarHitTestWindows[$hwnd.ToInt64().ToString()]
                if (-not $target) { return [IntPtr]::Zero }
                $raw = $lParam.ToInt64()
                $screenX = [int]($raw -band 0xffff); if ($screenX -ge 0x8000) { $screenX -= 0x10000 }
                $screenY = [int](($raw -shr 16) -band 0xffff); if ($screenY -ge 0x8000) { $screenY -= 0x10000 }
                $pt = $target.PointFromScreen([System.Windows.Point]::new($screenX, $screenY))
                $titleBarH = Get-TitleBarDragHeight -Window $target
                if ($pt.X -lt 0 -or $pt.X -gt $target.ActualWidth) { return [IntPtr]::Zero }
                if ($pt.Y -lt 4 -or $pt.Y -gt $titleBarH) { return [IntPtr]::Zero }
                if (Test-IsWindowCommandPoint -Window $target -Point $pt) { return [IntPtr]::Zero }
                $handled.Value = $true
                return [IntPtr]$HTCAPTION
            } catch { return [IntPtr]::Zero }
        }
        $script:TitleBarHitTestHooks[$key] = $hook
        $source.AddHook($hook)
    } catch { $null = $_ }
}

function Remove-NativeTitleBarHitTestHook {
    param([MahApps.Metro.Controls.MetroWindow]$Window)
    try {
        $helper = [System.Windows.Interop.WindowInteropHelper]::new($Window)
        $key = $helper.Handle.ToInt64().ToString()
        if ($script:TitleBarHitTestHooks.ContainsKey($key)) {
            $source = [System.Windows.Interop.HwndSource]::FromHwnd($helper.Handle)
            if ($source) { $source.RemoveHook($script:TitleBarHitTestHooks[$key]) }
            $script:TitleBarHitTestHooks.Remove($key)
        }
        if ($script:TitleBarHitTestWindows.ContainsKey($key)) {
            $script:TitleBarHitTestWindows.Remove($key)
        }
    } catch { $null = $_ }
}

function Install-TitleBarDragFallback {
    param([MahApps.Metro.Controls.MetroWindow]$Window)
    $Window.Add_SourceInitialized({ param($s, $e) Add-NativeTitleBarHitTestHook -Window $s })
    $Window.Add_Closed({ param($s, $e) Remove-NativeTitleBarHitTestHook -Window $s })
    $Window.Add_PreviewMouseLeftButtonDown({
        param($s, $e)
        try {
            if ($s.WindowState -eq [System.Windows.WindowState]::Maximized) { return }
            $titleBarH = Get-TitleBarDragHeight -Window $s
            $pos = $e.GetPosition($s)
            if ($pos.Y -lt 4 -or $pos.Y -gt $titleBarH) { return }
            if (Test-IsWindowCommandPoint -Window $s -Point $pos) { return }
            foreach ($ancestor in Get-InputAncestors -Start ($e.OriginalSource -as [System.Windows.DependencyObject])) {
                if ($ancestor -is [System.Windows.Controls.Primitives.ButtonBase]) { return }
            }
            $s.DragMove()
            $e.Handled = $true
        } catch { $null = $_ }
    })
}

# ---------------------------------------------------------------------------
# Module state (populated lazily on first show)
# ---------------------------------------------------------------------------

$script:DTRoot      = $null
$script:DTContext   = @{}
$script:IARoot      = $null
$script:IAContext   = @{}
$script:HistoryItems = New-Object 'System.Collections.ObjectModel.ObservableCollection[psobject]'
$script:IAItems      = New-Object 'System.Collections.ObjectModel.ObservableCollection[psobject]'

# ---------------------------------------------------------------------------
# DETECTION TESTER MODULE - handler scriptblocks
# ---------------------------------------------------------------------------

$script:DTOnTypeChanged = {
    $ctx = $script:DTContext
    if (-not $ctx.cboType) { return }
    $sel = Get-DTComboValue -Combo $ctx.cboType
    $panels = @{
        'RegistryKeyValue'         = $ctx.panelRegistryKeyValue
        'RegistryKey'              = $ctx.panelRegistryKey
        'File'                     = $ctx.panelFile
        'Script'                   = $ctx.panelScript
        'Compound (import-only)'   = $ctx.panelCompound
    }
    foreach ($k in $panels.Keys) {
        if ($null -ne $panels[$k]) {
            $panels[$k].Visibility = if ($k -eq $sel) { 'Visible' } else { 'Collapsed' }
        }
    }
}

$script:DTGatherClause = {
    $ctx  = $script:DTContext
    $type = Get-DTComboValue -Combo $ctx.cboType
    switch ($type) {
        'RegistryKeyValue' {
            return @{
                Type                = 'RegistryKeyValue'
                Hive                = if ($ctx.rkvHiveHKCU.IsChecked) { 'CurrentUser' } else { 'LocalMachine' }
                Is64Bit             = [bool]$ctx.rkvView64.IsChecked
                RegistryKeyRelative = $ctx.rkvKeyPath.Text
                ValueName           = $ctx.rkvValueName.Text
                Operator            = Get-DTComboValue -Combo $ctx.rkvOperator
                ExpectedValue       = $ctx.rkvExpected.Text
                PropertyType        = 'String'
            }
        }
        'RegistryKey' {
            return @{
                Type                = 'RegistryKey'
                Hive                = if ($ctx.rkHiveHKCU.IsChecked) { 'CurrentUser' } else { 'LocalMachine' }
                Is64Bit             = [bool]$ctx.rkView64.IsChecked
                RegistryKeyRelative = $ctx.rkKeyPath.Text
            }
        }
        'File' {
            return @{
                Type          = 'File'
                FilePath      = $ctx.fileFilePath.Text
                FileName      = $ctx.fileFileName.Text
                PropertyType  = Get-DTComboValue -Combo $ctx.filePropType
                Operator      = Get-DTComboValue -Combo $ctx.fileOperator
                ExpectedValue = $ctx.fileExpected.Text
                Is64Bit       = [bool]$ctx.fileView64.IsChecked
            }
        }
        'Script' {
            return @{
                Type           = 'Script'
                ScriptText     = $ctx.scriptText.Text
                ScriptLanguage = 'PowerShell'
            }
        }
        'Compound (import-only)' {
            if ($null -eq $ctx.CompoundClauses -or $ctx.CompoundClauses.Count -eq 0) {
                return $null
            }
            return @{
                Type      = 'Compound'
                Connector = $ctx.compoundConnector.Text
                Clauses   = $ctx.CompoundClauses
            }
        }
    }
}

$script:DTValidateClause = {
    param([hashtable]$Clause)
    if ($null -eq $Clause) { return 'No clause to evaluate.' }
    switch ($Clause.Type) {
        'RegistryKeyValue' {
            if ([string]::IsNullOrWhiteSpace($Clause.RegistryKeyRelative)) { return 'Key path is required.' }
            if ([string]::IsNullOrWhiteSpace($Clause.ValueName))           { return 'Value name is required.' }
        }
        'RegistryKey' {
            if ([string]::IsNullOrWhiteSpace($Clause.RegistryKeyRelative)) { return 'Key path is required.' }
        }
        'File' {
            if ([string]::IsNullOrWhiteSpace($Clause.FilePath)) { return 'File path is required.' }
            if ([string]::IsNullOrWhiteSpace($Clause.FileName)) { return 'File name is required.' }
            if ($Clause.PropertyType -eq 'Version' -and [string]::IsNullOrWhiteSpace($Clause.ExpectedValue)) {
                return 'Expected version is required when Property type = Version.'
            }
        }
        'Script' {
            if ([string]::IsNullOrWhiteSpace($Clause.ScriptText)) { return 'Script body is required.' }
        }
        'Compound' {
            if (-not $Clause.Clauses -or $Clause.Clauses.Count -eq 0) { return 'Compound clause has no children (import a manifest with Clauses).' }
        }
    }
    return $null
}

$script:DTPostTestResult = {
    param([hashtable]$Result, [bool]$Negative, [hashtable]$Clause)
    $ctx = $script:DTContext
    # Apply negative-test mode: pass becomes the inverse of detected.
    $effectivePass = if ($Negative) { -not $Result.Detected } else { [bool]$Result.Detected }
    $ctx.txtResultBadge.Text   = if ($effectivePass) { 'Pass' } else { 'Fail' }
    $ctx.txtResultDetails.Text = $Result.Details

    $row = [PSCustomObject]@{
        Time     = (Get-Date -Format 'HH:mm:ss')
        Type     = $Result.Type
        Mode     = if ($Negative) { 'Negative' } else { 'Positive' }
        Target   = [string]$Result.Target
        Expected = [string]$Result.ExpectedValue
        Found    = [string]$Result.ActualValue
        Result   = if ($effectivePass) { 'Pass' } else { 'Fail' }
        # Snapshot the full clause + negative flag so "Use selected" can
        # restore the input panel exactly, even after the underlying app
        # has been uninstalled.
        Clause   = $Clause
        Negative = $Negative
    }
    $script:HistoryItems.Insert(0, $row)
    Write-Log ("Test: type={0} mode={1} result={2} target='{3}'" -f $Result.Type, $row.Mode, $row.Result, $row.Target)
}

$script:DTRunTest = {
    $ctx = $script:DTContext
    $clause = & $script:DTGatherClause
    $validationError = & $script:DTValidateClause -Clause $clause
    if ($validationError) {
        $ctx.txtResultBadge.Text   = 'Invalid clause'
        $ctx.txtResultDetails.Text = $validationError
        return
    }

    $negative = [bool]$ctx.chkNegativeMode.IsChecked

    if ($clause.Type -eq 'Script') {
        # Async path: scripts can take seconds. Run on a background runspace
        # and poll via DispatcherTimer so the UI stays responsive.
        $ctx.btnTest.IsEnabled     = $false
        $ctx.txtResultBadge.Text   = 'Running...'
        $ctx.txtResultDetails.Text = 'Script test in flight (30s timeout). UI stays responsive.'

        $modulePath = Join-Path $PSScriptRoot 'Module\DetectionTesterCommon.psd1'

        $rs = [runspacefactory]::CreateRunspace()
        $rs.ApartmentState = 'STA'
        $rs.Open()
        $ps = [powershell]::Create()
        $ps.Runspace = $rs
        [void]$ps.AddScript({
            param($mp, $scriptText)
            Import-Module $mp -Force -DisableNameChecking
            Test-ScriptDetection -ScriptText $scriptText -TimeoutSeconds 30
        })
        [void]$ps.AddArgument($modulePath)
        [void]$ps.AddArgument([string]$clause.ScriptText)

        $async = $ps.BeginInvoke()

        # Capture locals for the timer Tick closure.
        $timer       = New-Object System.Windows.Threading.DispatcherTimer
        $timer.Interval = [TimeSpan]::FromMilliseconds(200)
        $localCtx       = $ctx
        $localClause    = $clause
        $localNegative  = $negative
        $localPs        = $ps
        $localRs        = $rs
        $localAsync     = $async
        $postResult     = $script:DTPostTestResult

        $timer.Add_Tick({
            if (-not $localAsync.IsCompleted) { return }
            $timer.Stop()
            try {
                $output = $localPs.EndInvoke($localAsync)
                $resultObj = $output | Where-Object { $_ -is [hashtable] } | Select-Object -Last 1
                if ($null -eq $resultObj) {
                    $localCtx.txtResultBadge.Text   = 'Error'
                    $localCtx.txtResultDetails.Text = 'Script test returned no result hashtable.'
                } else {
                    & $postResult -Result $resultObj -Negative $localNegative -Clause $localClause
                }
            } catch {
                $localCtx.txtResultBadge.Text   = 'Error'
                $localCtx.txtResultDetails.Text = "Async script test failed: $($_.Exception.Message)"
                Write-Log "Script async path error: $($_.Exception.Message)" -Level ERROR -Quiet
            } finally {
                # Cleanup failures aren't actionable; log and continue so the
                # button re-enables even if the runspace is in a weird state.
                try { $localPs.Dispose() } catch { Write-Log "DTRunTest async: ps.Dispose swallowed: $($_.Exception.Message)" -Level WARN -Quiet }
                try { $localRs.Close(); $localRs.Dispose() } catch { Write-Log "DTRunTest async: rs cleanup swallowed: $($_.Exception.Message)" -Level WARN -Quiet }
                $localCtx.btnTest.IsEnabled = $true
            }
        }.GetNewClosure())
        $timer.Start()
        return
    }

    # Sync path for non-Script types (registry/file/compound are fast).
    $result = switch ($clause.Type) {
        'RegistryKeyValue' { Test-RegistryKeyValueDetection -RegistryKeyRelative $clause.RegistryKeyRelative -ValueName $clause.ValueName -ExpectedValue $clause.ExpectedValue -Operator $clause.Operator -PropertyType $clause.PropertyType -Is64Bit $clause.Is64Bit -Hive $clause.Hive }
        'RegistryKey'      { Test-RegistryKeyDetection      -RegistryKeyRelative $clause.RegistryKeyRelative -Is64Bit $clause.Is64Bit -Hive $clause.Hive }
        'File'             { Test-FileDetection             -FilePath $clause.FilePath -FileName $clause.FileName -PropertyType $clause.PropertyType -ExpectedValue $clause.ExpectedValue -Operator $clause.Operator -Is64Bit $clause.Is64Bit }
        'Compound'         { Test-CompoundDetection         -Connector $clause.Connector -Clauses $clause.Clauses }
    }
    & $script:DTPostTestResult -Result $result -Negative $negative -Clause $clause
}

$script:DTImportManifest = {
    $ctx = $script:DTContext
    $dlg = New-Object Microsoft.Win32.OpenFileDialog
    $dlg.Filter = 'Stage manifest (*.json)|*.json|All files (*.*)|*.*'
    $dlg.Title  = 'Import stage-manifest.json'
    if (-not $dlg.ShowDialog($window)) { return }

    try {
        $det = Import-DetectionManifest -Path $dlg.FileName
    } catch {
        # Brand: no MessageBox::Show; surface errors in the result panel where the user is already looking.
        $ctx.txtResultBadge.Text   = 'Import failed'
        $ctx.txtResultDetails.Text = "Failed to import '$($dlg.FileName)': $($_.Exception.Message)"
        Write-Log "Import-DetectionManifest failed: $($_.Exception.Message)" -Level WARN -Quiet
        return
    }

    if ($det.Type -eq 'Compound') {
        $ctx.cboType.SelectedIndex = 4
        $ctx.compoundConnector.Text = [string]$det.Connector
        $ctx.compoundClauses.Items.Clear()
        foreach ($c in $det.Clauses) {
            $line = switch ($c.Type) {
                'RegistryKeyValue' { "RKV:  $($c.Hive)\$($c.RegistryKeyRelative) [$($c.ValueName)] $($c.Operator) '$($c.ExpectedValue)'" }
                'RegistryKey'      { "RK:   $($c.Hive)\$($c.RegistryKeyRelative)" }
                'File'             { "FILE: $($c.FilePath)\$($c.FileName) ($($c.PropertyType) $($c.Operator) '$($c.ExpectedValue)')" }
                'Script'           { "SCRIPT: $($c.ScriptText.Substring(0,[Math]::Min(60,$c.ScriptText.Length)))..." }
                default            { "?: $($c.Type)" }
            }
            [void]$ctx.compoundClauses.Items.Add($line)
        }
        $ctx.CompoundClauses = $det.Clauses
    } elseif ($det.Type -eq 'RegistryKeyValue') {
        $ctx.cboType.SelectedIndex = 0
        $ctx.rkvKeyPath.Text   = [string]$det.RegistryKeyRelative
        $ctx.rkvValueName.Text = [string]$det.ValueName
        $ctx.rkvExpected.Text  = [string]$det.ExpectedValue
        $ctx.rkvHiveHKCU.IsChecked = ($det.Hive -eq 'CurrentUser')
        $ctx.rkvHiveHKLM.IsChecked = ($det.Hive -ne 'CurrentUser')
        $ctx.rkvView64.IsChecked   = [bool]$det.Is64Bit
        $ctx.rkvView32.IsChecked   = -not [bool]$det.Is64Bit
        # Find operator combo by content
        for ($i = 0; $i -lt $ctx.rkvOperator.Items.Count; $i++) {
            if ([string]$ctx.rkvOperator.Items[$i].Content -eq [string]$det.Operator) { $ctx.rkvOperator.SelectedIndex = $i; break }
        }
    } elseif ($det.Type -eq 'RegistryKey') {
        $ctx.cboType.SelectedIndex = 1
        $ctx.rkKeyPath.Text = [string]$det.RegistryKeyRelative
        $ctx.rkHiveHKCU.IsChecked = ($det.Hive -eq 'CurrentUser')
        $ctx.rkHiveHKLM.IsChecked = ($det.Hive -ne 'CurrentUser')
        $ctx.rkView64.IsChecked   = [bool]$det.Is64Bit
        $ctx.rkView32.IsChecked   = -not [bool]$det.Is64Bit
    } elseif ($det.Type -eq 'File') {
        $ctx.cboType.SelectedIndex = 2
        $ctx.fileFilePath.Text = [string]$det.FilePath
        $ctx.fileFileName.Text = [string]$det.FileName
        $ctx.fileExpected.Text = [string]$det.ExpectedValue
        $ctx.fileView64.IsChecked = [bool]$det.Is64Bit
        $ctx.fileView32.IsChecked = -not [bool]$det.Is64Bit
        for ($i = 0; $i -lt $ctx.filePropType.Items.Count; $i++) {
            if ([string]$ctx.filePropType.Items[$i].Content -eq [string]$det.PropertyType) { $ctx.filePropType.SelectedIndex = $i; break }
        }
        for ($i = 0; $i -lt $ctx.fileOperator.Items.Count; $i++) {
            if ([string]$ctx.fileOperator.Items[$i].Content -eq [string]$det.Operator) { $ctx.fileOperator.SelectedIndex = $i; break }
        }
    } elseif ($det.Type -eq 'Script') {
        $ctx.cboType.SelectedIndex = 3
        $ctx.scriptText.Text = [string]$det.ScriptText
    }

    Write-Log ("Imported manifest: type={0} from {1}" -f $det.Type, $dlg.FileName)
}

$script:DTRepopulateFromHistory = {
    $ctx = $script:DTContext
    $row = $ctx.dgHistory.SelectedItem
    if ($null -eq $row) { return }
    $clause = $row.Clause
    if ($null -eq $clause) {
        # Older history row without a snapshot; only switch the type.
        switch ($row.Type) {
            'RegistryKeyValue' { $ctx.cboType.SelectedIndex = 0 }
            'RegistryKey'      { $ctx.cboType.SelectedIndex = 1 }
            'File'             { $ctx.cboType.SelectedIndex = 2 }
            'Script'           { $ctx.cboType.SelectedIndex = 3 }
            'Compound'         { $ctx.cboType.SelectedIndex = 4 }
        }
        return
    }

    $ctx.chkNegativeMode.IsChecked = [bool]$row.Negative

    switch ($clause.Type) {
        'RegistryKeyValue' {
            $ctx.cboType.SelectedIndex = 0
            $ctx.rkvKeyPath.Text       = [string]$clause.RegistryKeyRelative
            $ctx.rkvValueName.Text     = [string]$clause.ValueName
            $ctx.rkvExpected.Text      = [string]$clause.ExpectedValue
            $ctx.rkvHiveHKCU.IsChecked = ($clause.Hive -eq 'CurrentUser')
            $ctx.rkvHiveHKLM.IsChecked = ($clause.Hive -ne 'CurrentUser')
            $ctx.rkvView64.IsChecked   = [bool]$clause.Is64Bit
            $ctx.rkvView32.IsChecked   = -not [bool]$clause.Is64Bit
            for ($i = 0; $i -lt $ctx.rkvOperator.Items.Count; $i++) {
                if ([string]$ctx.rkvOperator.Items[$i].Content -eq [string]$clause.Operator) { $ctx.rkvOperator.SelectedIndex = $i; break }
            }
        }
        'RegistryKey' {
            $ctx.cboType.SelectedIndex = 1
            $ctx.rkKeyPath.Text        = [string]$clause.RegistryKeyRelative
            $ctx.rkHiveHKCU.IsChecked  = ($clause.Hive -eq 'CurrentUser')
            $ctx.rkHiveHKLM.IsChecked  = ($clause.Hive -ne 'CurrentUser')
            $ctx.rkView64.IsChecked    = [bool]$clause.Is64Bit
            $ctx.rkView32.IsChecked    = -not [bool]$clause.Is64Bit
        }
        'File' {
            $ctx.cboType.SelectedIndex = 2
            $ctx.fileFilePath.Text     = [string]$clause.FilePath
            $ctx.fileFileName.Text     = [string]$clause.FileName
            $ctx.fileExpected.Text     = [string]$clause.ExpectedValue
            $ctx.fileView64.IsChecked  = [bool]$clause.Is64Bit
            $ctx.fileView32.IsChecked  = -not [bool]$clause.Is64Bit
            for ($i = 0; $i -lt $ctx.filePropType.Items.Count; $i++) {
                if ([string]$ctx.filePropType.Items[$i].Content -eq [string]$clause.PropertyType) { $ctx.filePropType.SelectedIndex = $i; break }
            }
            for ($i = 0; $i -lt $ctx.fileOperator.Items.Count; $i++) {
                if ([string]$ctx.fileOperator.Items[$i].Content -eq [string]$clause.Operator) { $ctx.fileOperator.SelectedIndex = $i; break }
            }
        }
        'Script' {
            $ctx.cboType.SelectedIndex = 3
            $ctx.scriptText.Text       = [string]$clause.ScriptText
        }
        'Compound' {
            $ctx.cboType.SelectedIndex = 4
            $ctx.compoundConnector.Text = [string]$clause.Connector
            $ctx.compoundClauses.Items.Clear()
            foreach ($c in $clause.Clauses) {
                [void]$ctx.compoundClauses.Items.Add(("{0}: {1}" -f $c.Type, ($c | ConvertTo-Json -Compress)))
            }
            $ctx.CompoundClauses = $clause.Clauses
        }
    }
    Write-Log ("Repopulate: restored full clause from history (type={0}, negative={1})" -f $clause.Type, $row.Negative)
}

# Helper: convert ObservableCollection of PSCustomObject to DataTable for export
function ConvertTo-DTHistoryDataTable {
    param([Parameter(Mandatory)]$Items)
    $dt = New-Object System.Data.DataTable
    foreach ($col in @('Time','Type','Mode','Target','Expected','Found','Result')) {
        [void]$dt.Columns.Add($col, [string])
    }
    foreach ($r in $Items) {
        $row = $dt.NewRow()
        $row['Time']     = [string]$r.Time
        $row['Type']     = [string]$r.Type
        $row['Mode']     = [string]$r.Mode
        $row['Target']   = [string]$r.Target
        $row['Expected'] = [string]$r.Expected
        $row['Found']    = [string]$r.Found
        $row['Result']   = [string]$r.Result
        [void]$dt.Rows.Add($row)
    }
    return $dt
}

$script:DTExportHistoryCsv = {
    if ($script:HistoryItems.Count -eq 0) { return }
    $dlg = New-Object Microsoft.Win32.SaveFileDialog
    $dlg.Filter = 'CSV (*.csv)|*.csv'
    $dlg.InitialDirectory = $reportsDir
    $dlg.FileName = ('detection-history-{0}.csv' -f (Get-Date -Format 'yyyyMMdd-HHmmss'))
    if (-not $dlg.ShowDialog($window)) { return }
    $dt = ConvertTo-DTHistoryDataTable -Items $script:HistoryItems
    Export-DetectionResultsCsv -DataTable $dt -OutputPath $dlg.FileName
}

$script:DTExportHistoryHtml = {
    if ($script:HistoryItems.Count -eq 0) { return }
    $dlg = New-Object Microsoft.Win32.SaveFileDialog
    $dlg.Filter = 'HTML (*.html)|*.html'
    $dlg.InitialDirectory = $reportsDir
    $dlg.FileName = ('detection-history-{0}.html' -f (Get-Date -Format 'yyyyMMdd-HHmmss'))
    if (-not $dlg.ShowDialog($window)) { return }
    $dt = ConvertTo-DTHistoryDataTable -Items $script:HistoryItems
    Export-DetectionResultsHtml -DataTable $dt -OutputPath $dlg.FileName
}

$script:DTClearHistory = {
    $script:HistoryItems.Clear()
}

# ---------------------------------------------------------------------------
# INSTALLED APPS MODULE - handler scriptblocks
# ---------------------------------------------------------------------------

$script:IARefresh = {
    $apps = Get-InstalledApplications
    $script:IAContext.AllApps = $apps
    & $script:IAApplyFilter
}

$script:IAApplyFilter = {
    $ctx = $script:IAContext
    if ($null -eq $ctx.AllApps) { return }
    $filterText = [string]$ctx.txtFilter.Text
    $scopeSel   = Get-DTComboValue -Combo $ctx.cboScope
    $script:IAItems.Clear()
    foreach ($a in $ctx.AllApps) {
        if ($filterText) {
            $hay = "$($a.DisplayName) $($a.Publisher) $($a.DisplayVersion)"
            if ($hay -notmatch [regex]::Escape($filterText)) { continue }
        }
        switch ($scopeSel) {
            'Machine (HKLM)' { if ($a.Scope -ne 'Machine') { continue } }
            'User (HKCU)'    { if ($a.Scope -ne 'User')    { continue } }
        }
        $script:IAItems.Add($a)
    }
}

$script:IAOnAppSelected = {
    $ctx = $script:IAContext
    $a = $ctx.dgApps.SelectedItem
    $hasSelection = ($null -ne $a)
    # Brand: hide action buttons when not applicable, don't grey-disable.
    $vis = if ($hasSelection) { 'Visible' } else { 'Collapsed' }
    $ctx.btnUseForDetection.Visibility = $vis
    $ctx.btnCopyDetails.Visibility     = $vis
    if (-not $hasSelection) {
        foreach ($f in @('dtlDisplayName','dtlPublisher','dtlDisplayVersion','dtlArchScope','dtlRegistryKey','dtlUninstallString','dtlQuietUninstall','dtlInstallLocation','dtlInstallDate')) {
            if ($ctx[$f]) { $ctx[$f].Text = '' }
        }
        return
    }
    $ctx.dtlDisplayName.Text     = [string]$a.DisplayName
    $ctx.dtlPublisher.Text       = [string]$a.Publisher
    $ctx.dtlDisplayVersion.Text  = [string]$a.DisplayVersion
    $ctx.dtlArchScope.Text       = "$($a.Architecture) / $($a.Scope)"
    $ctx.dtlRegistryKey.Text     = [string]$a.RegistryKey
    $ctx.dtlUninstallString.Text = [string]$a.UninstallString
    $ctx.dtlQuietUninstall.Text  = [string]$a.QuietUninstallString
    $ctx.dtlInstallLocation.Text = [string]$a.InstallLocation
    $ctx.dtlInstallDate.Text     = [string]$a.InstallDate
}

$script:IAUseForDetection = {
    $a = $script:IAContext.dgApps.SelectedItem
    if ($null -eq $a) { return }
    & $script:ShowModule 'DetectionTester'
    $dt = $script:DTContext
    if (-not $dt.cboType) { return }
    $dt.cboType.SelectedIndex = 0  # RegistryKeyValue
    $dt.rkvKeyPath.Text   = [string]$a.RegistryKeyRelative
    $dt.rkvValueName.Text = 'DisplayVersion'
    $dt.rkvExpected.Text  = [string]$a.DisplayVersion
    $dt.rkvHiveHKCU.IsChecked = ($a.Hive -eq 'CurrentUser')
    $dt.rkvHiveHKLM.IsChecked = ($a.Hive -ne 'CurrentUser')
    # Architecture: x64 -> Registry64 view; x86 -> Registry32 view (WOW6432Node already in path)
    # For HKCU there's no WOW split; default to 64.
    $dt.rkvView64.IsChecked = ($a.Architecture -ne 'x86')
    $dt.rkvView32.IsChecked = ($a.Architecture -eq 'x86')
    # Operator default: IsEquals
    $dt.rkvOperator.SelectedIndex = 0
    Write-Log ("Use for detection: app='{0}' key={1}" -f $a.DisplayName, $a.RegistryKey)
}

$script:IACopyDetails = {
    $a = $script:IAContext.dgApps.SelectedItem
    if ($null -eq $a) { return }
    $lines = @(
        ("DisplayName:          $($a.DisplayName)"),
        ("Publisher:            $($a.Publisher)"),
        ("DisplayVersion:       $($a.DisplayVersion)"),
        ("Architecture:         $($a.Architecture)"),
        ("Scope:                $($a.Scope)"),
        ("RegistryKey:          $($a.RegistryKey)"),
        ("UninstallString:      $($a.UninstallString)"),
        ("QuietUninstallString: $($a.QuietUninstallString)"),
        ("InstallLocation:      $($a.InstallLocation)"),
        ("InstallDate:          $($a.InstallDate)")
    ) -join "`r`n"
    [System.Windows.Clipboard]::SetText($lines)
}

# ---------------------------------------------------------------------------
# Module loaders
# ---------------------------------------------------------------------------

$script:LoadDetectionTester = {
    $root = Import-DTXaml -Path (Join-Path $PSScriptRoot 'Modules\DetectionTester.xaml')
    $names = @(
        'cboType','btnImportManifest','chkNegativeMode',
        'panelRegistryKeyValue','panelRegistryKey','panelFile','panelScript','panelCompound',
        'rkvHiveHKLM','rkvHiveHKCU','rkvView64','rkvView32','rkvKeyPath','rkvValueName','rkvOperator','rkvExpected',
        'rkHiveHKLM','rkHiveHKCU','rkView64','rkView32','rkKeyPath',
        'fileView64','fileView32','fileFilePath','fileFileName','filePropType','fileOperator','fileExpected',
        'scriptText',
        'compoundConnector','compoundClauses',
        'btnTest','txtResultBadge','txtResultDetails',
        'dgHistory','btnRepopulate','btnExportCsv','btnExportHtml','btnClearHistory'
    )
    $ctx = @{ Root = $root }
    foreach ($n in $names) { $ctx[$n] = $root.FindName($n) }
    $script:DTContext = $ctx

    $ctx.dgHistory.ItemsSource = $script:HistoryItems

    # PS51-WPF-001: handlers attached inside this loader scriptblock can't
    # read $script:* directly through .GetNewClosure(). Capture each
    # script-scope scriptblock as a local variable; the closure picks up
    # the local via GetNewClosure's variable capture.
    $onTypeChanged   = $script:DTOnTypeChanged
    $runTest         = $script:DTRunTest
    $importManifest  = $script:DTImportManifest
    $repopulate      = $script:DTRepopulateFromHistory
    $exportCsv       = $script:DTExportHistoryCsv
    $exportHtml      = $script:DTExportHistoryHtml
    $clearHistory    = $script:DTClearHistory
    $localCtx        = $ctx

    $ctx.cboType.Add_SelectionChanged({ & $onTypeChanged }.GetNewClosure())
    $ctx.btnTest.Add_Click({ & $runTest }.GetNewClosure())
    $ctx.btnImportManifest.Add_Click({ & $importManifest }.GetNewClosure())
    $ctx.btnRepopulate.Add_Click({ & $repopulate }.GetNewClosure())
    $ctx.btnExportCsv.Add_Click({ & $exportCsv }.GetNewClosure())
    $ctx.btnExportHtml.Add_Click({ & $exportHtml }.GetNewClosure())
    $ctx.btnClearHistory.Add_Click({ & $clearHistory }.GetNewClosure())
    # Hide btnRepopulate until a history row is selected.
    $ctx.dgHistory.Add_SelectionChanged({
        $sel = ($null -ne $localCtx.dgHistory.SelectedItem)
        $localCtx.btnRepopulate.Visibility = if ($sel) { 'Visible' } else { 'Collapsed' }
    }.GetNewClosure())

    & $onTypeChanged
    return $root
}

$script:LoadInstalledApps = {
    $root = Import-DTXaml -Path (Join-Path $PSScriptRoot 'Modules\InstalledApps.xaml')
    $names = @(
        'txtFilter','cboScope','btnRefresh','btnUseForDetection','btnCopyDetails',
        'dgApps',
        'dtlDisplayName','dtlPublisher','dtlDisplayVersion','dtlArchScope','dtlRegistryKey',
        'dtlUninstallString','dtlQuietUninstall','dtlInstallLocation','dtlInstallDate'
    )
    $ctx = @{ Root = $root }
    foreach ($n in $names) { $ctx[$n] = $root.FindName($n) }
    $script:IAContext = $ctx

    $ctx.dgApps.ItemsSource = $script:IAItems

    # PS51-WPF-001: capture script-scope scriptblocks as locals so the
    # GetNewClosure'd handlers can resolve them when the dispatcher fires.
    $refresh         = $script:IARefresh
    $useForDetection = $script:IAUseForDetection
    $copyDetails     = $script:IACopyDetails
    $applyFilter     = $script:IAApplyFilter
    $onAppSelected   = $script:IAOnAppSelected

    $ctx.btnRefresh.Add_Click({ & $refresh }.GetNewClosure())
    $ctx.btnUseForDetection.Add_Click({ & $useForDetection }.GetNewClosure())
    $ctx.btnCopyDetails.Add_Click({ & $copyDetails }.GetNewClosure())
    $ctx.txtFilter.Add_TextChanged({ & $applyFilter }.GetNewClosure())
    $ctx.cboScope.Add_SelectionChanged({ & $applyFilter }.GetNewClosure())
    $ctx.dgApps.Add_SelectionChanged({ & $onAppSelected }.GetNewClosure())

    # First-time: enumerate ARP
    & $refresh
    return $root
}

# ---------------------------------------------------------------------------
# Prefs
# ---------------------------------------------------------------------------

$script:Prefs = Get-DTPrefs

# ---------------------------------------------------------------------------
# Load shell XAML
# ---------------------------------------------------------------------------

$xamlPath = Join-Path $PSScriptRoot 'MainWindow.xaml'
$window   = Import-DTXaml -Path $xamlPath
Install-TitleBarDragFallback -Window $window

# ---------------------------------------------------------------------------
# Wired controls
# ---------------------------------------------------------------------------

$txtAppVersion     = $window.FindName('txtAppVersion')
$txtModuleTitle    = $window.FindName('txtModuleTitle')
$txtModuleSubtitle = $window.FindName('txtModuleSubtitle')
$contentHost       = $window.FindName('contentHost')
$txtStatus         = $window.FindName('txtStatus')
$lblLogOutput      = $window.FindName('lblLogOutput')
$btnDetection      = $window.FindName('btnDetection')
$btnInstalled      = $window.FindName('btnInstalled')
$btnOptions        = $window.FindName('btnOptions')
$toggleTheme       = $window.FindName('toggleTheme')
$txtThemeLabel     = $window.FindName('txtThemeLabel')

$txtAppVersion.Text = 'v1.0.0'

# ---------------------------------------------------------------------------
# Theme runtime brushes (XAML literal SolidColorBrush values don't flip on
# MahApps ChangeTheme - must be set per-theme at runtime).
# See reference_srl_wpf_brand.md S6.
# ---------------------------------------------------------------------------

$script:DarkButtonBg      = [Windows.Media.BrushConverter]::new().ConvertFrom('#1E1E1E')
$script:DarkButtonBorder  = [Windows.Media.BrushConverter]::new().ConvertFrom('#555555')
$script:LightButtonBg     = [Windows.Media.BrushConverter]::new().ConvertFrom('#0078D4')
$script:LightButtonBorder = [Windows.Media.BrushConverter]::new().ConvertFrom('#006CBE')

$script:TitleBarBlue         = [Windows.Media.BrushConverter]::new().ConvertFrom('#0078D4')
$script:TitleBarBlueInactive = [Windows.Media.BrushConverter]::new().ConvertFrom('#4BA3E0')

$script:SidebarButtons = @('btnDetection','btnInstalled','btnOptions') |
    ForEach-Object { $window.FindName($_) } |
    Where-Object { $_ }

# ---------------------------------------------------------------------------
# Handler-reachable scriptblocks (shell)
# ---------------------------------------------------------------------------

$script:ApplyButtonTheme = {
    param([bool]$IsDark)
    $bg     = if ($IsDark) { $script:DarkButtonBg }     else { $script:LightButtonBg }
    $border = if ($IsDark) { $script:DarkButtonBorder } else { $script:LightButtonBorder }
    foreach ($b in $script:SidebarButtons) {
        $b.Background  = $bg
        $b.BorderBrush = $border
    }
    if ($IsDark) {
        $window.ClearValue([MahApps.Metro.Controls.MetroWindow]::WindowTitleBrushProperty)
        $window.ClearValue([MahApps.Metro.Controls.MetroWindow]::NonActiveWindowTitleBrushProperty)
    } else {
        $window.WindowTitleBrush          = $script:TitleBarBlue
        $window.NonActiveWindowTitleBrush = $script:TitleBarBlueInactive
    }
}

$script:ApplyTheme = {
    param([bool]$IsDark)
    $themeName = if ($IsDark) { 'Dark.Steel' } else { 'Light.Blue' }
    [ControlzEx.Theming.ThemeManager]::Current.ChangeTheme($window, $themeName) | Out-Null
    $txtThemeLabel.Text = if ($IsDark) { 'Dark Theme' } else { 'Light Theme' }
    if ($lblLogOutput) {
        $hex = if ($IsDark) { '#B0B0B0' } else { '#595959' }
        $lblLogOutput.Foreground = [Windows.Media.BrushConverter]::new().ConvertFrom($hex)
    }
    & $script:ApplyButtonTheme -IsDark $IsDark
    $script:Prefs.DarkMode = $IsDark
    Write-Log ("Theme applied: {0}" -f $themeName)
}

$script:ShowModule = {
    param([Parameter(Mandatory)][string]$Name)
    switch ($Name) {
        'DetectionTester' {
            if ($null -eq $script:DTRoot) {
                $script:DTRoot = & $script:LoadDetectionTester
            }
            $contentHost.Content    = $script:DTRoot
            $txtModuleTitle.Text    = 'Detection Tester'
            $txtModuleSubtitle.Text = 'Author and test MECM detection methods locally.'
            $txtStatus.Text         = 'Detection Tester selected.'
        }
        'InstalledApps' {
            if ($null -eq $script:IARoot) {
                $script:IARoot = & $script:LoadInstalledApps
            }
            $contentHost.Content    = $script:IARoot
            $txtModuleTitle.Text    = 'Installed Applications'
            $txtModuleSubtitle.Text = 'Browse ARP entries (HKLM and HKCU, x64 and x86).'
            $txtStatus.Text         = 'Installed Applications selected.'
        }
        default {
            & $script:ShowModule 'DetectionTester'
            return
        }
    }
    $script:Prefs.ActiveModule = $Name
}

$script:OnClosing = {
    Save-DTPrefs       -Prefs $script:Prefs
    Save-DTWindowState
    Write-Log 'Detection Tester WPF shell closed.'
}

# ---------------------------------------------------------------------------
# Options dialog (in-app MetroWindow per brand spec S22)
# ---------------------------------------------------------------------------

$script:ShowOptionsDialog = {
    $dlgXaml = @'
<Controls:MetroWindow
    xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
    xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
    xmlns:Controls="clr-namespace:MahApps.Metro.Controls;assembly=MahApps.Metro"
    Title="Options"
    Width="640" Height="440"
    MinWidth="560" MinHeight="380"
    WindowStartupLocation="CenterOwner"
    TitleCharacterCasing="Normal"
    ShowIconOnTitleBar="False"
    ResizeMode="CanResize"
    GlowBrush="{DynamicResource MahApps.Brushes.Accent}"
    NonActiveGlowBrush="{DynamicResource MahApps.Brushes.Accent}"
    BorderThickness="1">
    <Window.Resources>
        <ResourceDictionary>
            <ResourceDictionary.MergedDictionaries>
                <ResourceDictionary Source="pack://application:,,,/MahApps.Metro;component/Styles/Controls.xaml"/>
                <ResourceDictionary Source="pack://application:,,,/MahApps.Metro;component/Styles/Fonts.xaml"/>
                <ResourceDictionary Source="pack://application:,,,/MahApps.Metro;component/Styles/Themes/Dark.Steel.xaml"/>
            </ResourceDictionary.MergedDictionaries>
        </ResourceDictionary>
    </Window.Resources>
    <Grid>
        <Grid.ColumnDefinitions>
            <ColumnDefinition Width="200"/>
            <ColumnDefinition Width="1"/>
            <ColumnDefinition Width="*"/>
        </Grid.ColumnDefinitions>
        <Grid.RowDefinitions>
            <RowDefinition Height="*"/>
            <RowDefinition Height="Auto"/>
        </Grid.RowDefinitions>

        <!-- Left nav -->
        <Border Grid.Column="0" Grid.Row="0" Background="{DynamicResource MahApps.Brushes.Gray10}">
            <ListBox x:Name="lstNav" SelectedIndex="0" BorderThickness="0" Padding="0,8">
                <ListBoxItem Content="About"   Padding="14,8" FontSize="13"/>
                <ListBoxItem Content="Logging" Padding="14,8" FontSize="13"/>
            </ListBox>
        </Border>
        <Border Grid.Column="1" Grid.Row="0" Background="{DynamicResource MahApps.Brushes.Gray8}"/>

        <!-- Right pane -->
        <ScrollViewer Grid.Column="2" Grid.Row="0" Padding="20,16" VerticalScrollBarVisibility="Auto">
            <Grid>
                <!-- About -->
                <StackPanel x:Name="paneAbout" Visibility="Visible">
                    <TextBlock Text="Detection Tester" FontSize="18" FontWeight="SemiBold"/>
                    <TextBlock x:Name="txtAboutVersion" Text="v1.0.0" FontSize="11"
                               Foreground="{DynamicResource MahApps.Brushes.Gray1}" Margin="0,2,0,12"/>
                    <TextBlock TextWrapping="Wrap" Margin="0,0,0,12"
                               Text="Local GUI for testing MECM application detection methods (RegistryKeyValue, RegistryKey, File, Script, Compound) against the machine the tool is running on, without deploying through MECM. Browses ARP entries (HKLM and HKCU) for fast clause authoring."/>
                    <Grid Margin="0,4">
                        <Grid.ColumnDefinitions>
                            <ColumnDefinition Width="120"/>
                            <ColumnDefinition Width="*"/>
                        </Grid.ColumnDefinitions>
                        <Grid.RowDefinitions>
                            <RowDefinition Height="Auto"/>
                            <RowDefinition Height="Auto"/>
                            <RowDefinition Height="Auto"/>
                            <RowDefinition Height="Auto"/>
                        </Grid.RowDefinitions>
                        <TextBlock Grid.Row="0" Grid.Column="0" Text="Author:"   FontWeight="SemiBold" Margin="0,3"/>
                        <TextBlock Grid.Row="0" Grid.Column="1" Text="Jason Ulbright"                  Margin="0,3"/>
                        <TextBlock Grid.Row="1" Grid.Column="0" Text="License:"  FontWeight="SemiBold" Margin="0,3"/>
                        <TextBlock Grid.Row="1" Grid.Column="1" Text="MIT"                             Margin="0,3"/>
                        <TextBlock Grid.Row="2" Grid.Column="0" Text="Stack:"    FontWeight="SemiBold" Margin="0,3"/>
                        <TextBlock Grid.Row="2" Grid.Column="1" Text="PowerShell 5.1 + WPF (MahApps)"  Margin="0,3"/>
                        <TextBlock Grid.Row="3" Grid.Column="0" Text="Module:"   FontWeight="SemiBold" Margin="0,3"/>
                        <TextBlock Grid.Row="3" Grid.Column="1" x:Name="txtAboutModule"                Margin="0,3" FontFamily="Cascadia Code, Consolas, Courier New" FontSize="11"/>
                    </Grid>
                </StackPanel>

                <!-- Logging -->
                <StackPanel x:Name="paneLogging" Visibility="Collapsed">
                    <TextBlock Text="Logging" FontSize="18" FontWeight="SemiBold" Margin="0,0,0,4"/>
                    <TextBlock TextWrapping="Wrap" Margin="0,0,0,12"
                               Foreground="{DynamicResource MahApps.Brushes.Gray1}"
                               Text="A new log file is created per session under %LOCALAPPDATA%\DetectionTester\logs. Old logs accumulate; clear them manually as needed."/>
                    <TextBlock Text="Current session log:" FontWeight="SemiBold" Margin="0,4,0,2"/>
                    <TextBox x:Name="txtLogPath" IsReadOnly="True"
                             FontFamily="Cascadia Code, Consolas, Courier New" FontSize="11"
                             Margin="0,0,0,8"/>
                    <TextBlock Text="Log folder:" FontWeight="SemiBold" Margin="0,8,0,2"/>
                    <TextBox x:Name="txtLogDir"  IsReadOnly="True"
                             FontFamily="Cascadia Code, Consolas, Courier New" FontSize="11"
                             Margin="0,0,0,8"/>
                    <StackPanel Orientation="Horizontal" Margin="0,4,0,0">
                        <Button x:Name="btnOpenLogFolder" Content="Open log folder"
                                Style="{DynamicResource MahApps.Styles.Button.Square}"
                                Controls:ControlsHelper.ContentCharacterCasing="Normal"
                                Padding="10,4" MinWidth="120" Margin="0,0,8,0"/>
                        <Button x:Name="btnOpenAppData"   Content="Open app data folder"
                                Style="{DynamicResource MahApps.Styles.Button.Square}"
                                Controls:ControlsHelper.ContentCharacterCasing="Normal"
                                Padding="10,4" MinWidth="160"/>
                    </StackPanel>
                </StackPanel>
            </Grid>
        </ScrollViewer>

        <!-- Footer -->
        <Border Grid.Row="1" Grid.ColumnSpan="3" Padding="16,12"
                Background="{DynamicResource MahApps.Brushes.Gray10}">
            <StackPanel Orientation="Horizontal" HorizontalAlignment="Right">
                <Button x:Name="btnOK"     Content="OK"     Style="{DynamicResource MahApps.Styles.Button.Square.Accent}" Controls:ControlsHelper.ContentCharacterCasing="Normal" IsDefault="True" MinWidth="90" Height="32" Margin="0,0,8,0"/>
                <Button x:Name="btnCancel" Content="Cancel" Style="{DynamicResource MahApps.Styles.Button.Square}"        Controls:ControlsHelper.ContentCharacterCasing="Normal" IsCancel="True"  MinWidth="90" Height="32"/>
            </StackPanel>
        </Border>
    </Grid>
</Controls:MetroWindow>
'@

    $reader = [System.Xml.XmlReader]::Create((New-Object System.IO.StringReader $dlgXaml))
    $dlg = [Windows.Markup.XamlReader]::Load($reader)
    Install-TitleBarDragFallback -Window $dlg

    # Theme-propagate to dialog
    $currentTheme = [ControlzEx.Theming.ThemeManager]::Current.DetectTheme($window)
    if ($currentTheme) {
        [ControlzEx.Theming.ThemeManager]::Current.ChangeTheme($dlg, $currentTheme) | Out-Null
    }

    $dlg.Owner = $window

    $lstNav            = $dlg.FindName('lstNav')
    $paneAbout         = $dlg.FindName('paneAbout')
    $paneLogging       = $dlg.FindName('paneLogging')
    $txtAboutVersion   = $dlg.FindName('txtAboutVersion')
    $txtAboutModule    = $dlg.FindName('txtAboutModule')
    $txtLogPath        = $dlg.FindName('txtLogPath')
    $txtLogDir         = $dlg.FindName('txtLogDir')
    $btnOpenLogFolder  = $dlg.FindName('btnOpenLogFolder')
    $btnOpenAppData    = $dlg.FindName('btnOpenAppData')
    $btnOK             = $dlg.FindName('btnOK')
    $btnCancel         = $dlg.FindName('btnCancel')

    # Populate About
    $txtAboutVersion.Text = 'v1.0.0'
    $modVersion = (Get-Module DetectionTesterCommon | Select-Object -First 1).Version
    $txtAboutModule.Text  = "DetectionTesterCommon $modVersion"

    # Populate Logging
    $txtLogPath.Text = $logPath
    $txtLogDir.Text  = $logDir

    # Nav
    $lstNav.Add_SelectionChanged({
        $idx = $lstNav.SelectedIndex
        $paneAbout.Visibility   = if ($idx -eq 0) { 'Visible' } else { 'Collapsed' }
        $paneLogging.Visibility = if ($idx -eq 1) { 'Visible' } else { 'Collapsed' }
    }.GetNewClosure())

    # PS51-WPF-001: handlers attached inside this $script: scriptblock can't
    # read script-root $logDir / $appDataDir directly (the closure scope
    # chain doesn't extend back). Capture as locals so .GetNewClosure picks
    # them up.
    $localLogDir     = $logDir
    $localAppDataDir = $appDataDir

    # Logging buttons - open in Explorer (no admin needed; verb=open is default)
    $btnOpenLogFolder.Add_Click({
        Start-Process -FilePath 'explorer.exe' -ArgumentList @($localLogDir)
    }.GetNewClosure())
    $btnOpenAppData.Add_Click({
        Start-Process -FilePath 'explorer.exe' -ArgumentList @($localAppDataDir)
    }.GetNewClosure())

    # Footer
    $btnOK.Add_Click({ $dlg.DialogResult = $true; $dlg.Close() }.GetNewClosure())
    $btnCancel.Add_Click({ $dlg.DialogResult = $false; $dlg.Close() }.GetNewClosure())

    [void]$dlg.ShowDialog()
}

# ---------------------------------------------------------------------------
# Event handlers - closures invoke $script: scriptblocks via & only
# ---------------------------------------------------------------------------

$btnDetection.Add_Click({ & $script:ShowModule 'DetectionTester' }.GetNewClosure())
$btnInstalled.Add_Click({ & $script:ShowModule 'InstalledApps'   }.GetNewClosure())
$btnOptions.Add_Click({ & $script:ShowOptionsDialog }.GetNewClosure())

$toggleTheme.Add_Toggled({
    $newDark = [bool]$toggleTheme.IsOn
    if ($script:Prefs.DarkMode -eq $newDark) { return }
    & $script:ApplyTheme -IsDark $newDark
}.GetNewClosure())

$window.Add_Closing({ & $script:OnClosing }.GetNewClosure())

# ---------------------------------------------------------------------------
# Initial state
# ---------------------------------------------------------------------------

$toggleTheme.IsOn = $script:Prefs.DarkMode
& $script:ApplyTheme -IsDark $script:Prefs.DarkMode
Restore-DTWindowState
& $script:ShowModule $script:Prefs.ActiveModule

Write-Log ("Detection Tester WPF shell starting (theme={0}, module={1})" -f `
    $(if ($script:Prefs.DarkMode) { 'Dark' } else { 'Light' }), `
    $script:Prefs.ActiveModule)

# ---------------------------------------------------------------------------
# Show
# ---------------------------------------------------------------------------

$null = $window.ShowDialog()
