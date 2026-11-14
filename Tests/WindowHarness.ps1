# Headless window harness for the Detection Method Tester tests.
#
# A WPF window needs an STA thread. Invoke-WpfSta runs a script block on one,
# and the script block returns plain data (never a WPF object), because a WPF
# object cannot cross to the Pester thread. The windows are shown off screen at
# -32000 so layout, templates and triggers run for real.
#
# One STA runspace serves a whole test file. A runspace per call hangs: the
# theme manager keeps handlers of closed windows from earlier threads, and the
# next theme change waits on their dead dispatchers. Call Close-WpfSession in
# AfterAll.

$script:HarnessRoot = Split-Path -Parent $PSScriptRoot
. "$PSScriptRoot\WindowMeasure.ps1"

$script:HarnessSession = $null

# Window tests open a window only when SUITE_ALLOW_WINDOW_TESTS is 1.
function Test-WindowTestsAllowed {
    [CmdletBinding()]
    param(
        [string]$OptIn = $env:SUITE_ALLOW_WINDOW_TESTS
    )
    if ($OptIn -ne '1') { return [pscustomobject]@{ Allowed = $false; Reason = 'SUITE_ALLOW_WINDOW_TESTS is not 1' } }
    [pscustomobject]@{ Allowed = $true; Reason = '' }
}

# Script text that every STA runspace starts with: assemblies, SuiteCommon, and
# helper functions to open a themed window and measure it.
$script:HarnessBootstrap = {
    $ErrorActionPreference = 'Stop'
    Add-Type -AssemblyName PresentationFramework, PresentationCore, WindowsBase, System.Windows.Forms
    foreach ($dll in 'Microsoft.Xaml.Behaviors.dll', 'ControlzEx.dll', 'MahApps.Metro.dll') {
        [void][System.Reflection.Assembly]::LoadFrom((Join-Path $Root "Lib\$dll"))
    }
    Import-Module (Join-Path $Root 'Lib\SuiteCommon\SuiteCommon.psd1') -Global -DisableNameChecking

    # Opens a window from a XAML file, or from XAML text, off screen with the theme applied.
    function Open-TestWindow {
        param([string]$XamlPath, [string]$XamlText, [bool]$Dark = $true, [double]$Width = 1440, [double]$Height = 900)
        if (-not $XamlText) { $XamlText = Get-Content -LiteralPath $XamlPath -Raw -Encoding UTF8 }
        [xml]$xaml = $XamlText
        $reader = New-Object System.Xml.XmlNodeReader $xaml
        $w = [System.Windows.Markup.XamlReader]::Load($reader)
        $w.WindowStartupLocation = 'Manual'
        # SUITE_VISIBLE_WINDOWS=1 (the visible lab run in the labadmin session) shows the window on the desktop.
        $visible = ($env:SUITE_VISIBLE_WINDOWS -eq '1')
        $w.Left = $(if ($visible) { 80 } else { -32000 }); $w.Top = $(if ($visible) { 80 } else { -32000 })
        $w.ShowInTaskbar = $visible
        $w.Width = $Width; $w.Height = $Height
        [void][ControlzEx.Theming.ThemeManager]::Current.ChangeTheme($w, $(if ($Dark) { 'Dark.Steel' } else { 'Light.Blue' }))
        $w.Show()
        Sync-Layout $w
        # MetroWindow fades its root grid in; measure after the fade ends.
        $root = Get-VisualDescendants $w | Where-Object { $_ -is [System.Windows.FrameworkElement] -and $_.Name -eq 'RootGrid' } | Select-Object -First 1
        for ($i = 0; $root -and $i -lt 100 -and $root.Opacity -lt 0.999; $i++) { Start-Sleep -Milliseconds 50; Sync-Layout $w }
        if ($visible) { Start-Sleep -Milliseconds 600 }
        return $w
    }

    # Defines the handler functions and script blocks of the start script in this runspace and sets the window
    # variables that they read, as the shell does. The text comes from Get-ShellRuntimePart (TestHelpers.ps1).
    # The shell's own script blocks then run on the real controls: ApplyTheme, the loaders, the click handlers.
    # The script scope of a runspace without a script file is its global scope, so $script: and $global: agree.
    function Initialize-ShellRuntime {
        param($Window, [hashtable]$Parts, [string]$LogPath)
        Import-Module (Join-Path $Root 'Module\DetectionTesterCommon.psd1') -Force -DisableNameChecking
        Initialize-Logging -LogPath $LogPath
        $global:window = $Window
        foreach ($name in $Parts.Functions.Keys) { Set-Item -Path "function:global:$name" -Value ([scriptblock]::Create($Parts.Functions[$name])) }
        # A script block written as a literal in a dot-sourced script is bound to this session state. One from
        # [scriptblock]::Create is not, and inside a closure (GetNewClosure) it would read $script: variables of the closure's own module.
        foreach ($name in $Parts.Blocks.Keys) { . ([scriptblock]::Create('$global:' + $name + ' = {' + $Parts.Blocks[$name] + "`n}")) }
        foreach ($text in $Parts.Assignments) { & ([scriptblock]::Create(($text -replace '^\$script:', '$global:'))) }
        $global:contentHost = $Window.FindName('contentHost')
        $global:txtModuleTitle = $Window.FindName('txtModuleTitle')
        $global:txtModuleSubtitle = $Window.FindName('txtModuleSubtitle')
        $global:txtStatus = $Window.FindName('txtStatus')
        $global:lblLogOutput = $Window.FindName('lblLogOutput')
        $global:txtThemeLabel = $Window.FindName('txtThemeLabel')
        $global:Prefs = @{ DarkMode = $true; ActiveModule = 'DetectionTester' }
        $global:ActiveView = 'DetectionTester'
        # The theme context of the shell: the same call as in the start script, with the inactive title bar brush of MainWindow.xaml.
        $global:ViewButtons = @(
            @{ Name = 'DetectionTester'; Button = $Window.FindName('btnDetection') }
            @{ Name = 'InstalledApps'; Button = $Window.FindName('btnInstalled') }
        )
        $global:TitleBarInactiveBrush = $Window.FindResource('MahApps.Brushes.WindowTitle.NonActive')
        $suiteBrushes = @{ TitleBarInactive = $global:TitleBarInactiveBrush }
        Initialize-SuiteTheme -Window $Window -IsDarkGetter { [bool]$global:Prefs.DarkMode } -ActiveViewGetter { $global:ActiveView } -ViewButtons $global:ViewButtons -OptionsButton $Window.FindName('btnOptions') -LogLabel $Window.FindName('lblLogOutput') -Brushes $suiteBrushes
        $global:reportsDir = [System.IO.Path]::GetDirectoryName($LogPath)
        $global:DTRoot = $null; $global:IARoot = $null; $global:DTContext = @{}; $global:IAContext = @{}
        $global:HistoryItems = New-Object 'System.Collections.ObjectModel.ObservableCollection[psobject]'
        $global:IAItems = New-Object 'System.Collections.ObjectModel.ObservableCollection[psobject]'
    }

    function Invoke-Click {
        param($Button)
        $Button.RaiseEvent((New-Object System.Windows.RoutedEventArgs([System.Windows.Controls.Primitives.ButtonBase]::ClickEvent)))
    }

    # Pumps the dispatcher (timers and queued events run) until the condition is true or the time is up.
    function Wait-Until {
        param($Window, [scriptblock]$Condition, [int]$TimeoutSeconds = 30)
        $end = [DateTime]::UtcNow.AddSeconds($TimeoutSeconds)
        while ([DateTime]::UtcNow -lt $end) {
            Sync-Layout $Window
            if (& $Condition) { return $true }
            Start-Sleep -Milliseconds 100
        }
        return $false
    }

    # Loads the XAML of one module into the content host, as the shell does, and fills it with sample data so
    # that the grids, the result texts and the detail panel are laid out with real text.
    function Show-AppView {
        param($Window, [hashtable]$Config, [hashtable]$View)
        $module = $Config.Modules[$View.Module]
        if (-not $global:ModuleRoots) { $global:ModuleRoots = @{} }
        $key = '{0}|{1}' -f $Window.GetHashCode(), $View.Module
        if (-not $global:ModuleRoots.ContainsKey($key)) {
            [xml]$x = Get-Content -LiteralPath (Join-Path $Root $module.Xaml) -Raw -Encoding UTF8
            $global:ModuleRoots[$key] = [System.Windows.Markup.XamlReader]::Load((New-Object System.Xml.XmlNodeReader $x))
        }
        $moduleRoot = $global:ModuleRoots[$key]
        $Window.FindName($Config.ContentHost).Content = $moduleRoot
        $Window.FindName('txtModuleTitle').Text = $module.Title
        $Window.FindName('txtModuleSubtitle').Text = $module.Subtitle
        $Window.FindName('txtStatus').Text = $module.Status
        if ($View.Module -eq 'DetectionTester') {
            foreach ($p in $Config.Panels) { $moduleRoot.FindName($p).Visibility = if ($p -eq $View.Panel) { 'Visible' } else { 'Collapsed' } }
            $moduleRoot.FindName('cboType').SelectedIndex = $View.TypeIndex
            $history = New-Object 'System.Collections.ObjectModel.ObservableCollection[psobject]'
            foreach ($r in $Config.SampleHistory) { $history.Add([PSCustomObject]$r) }
            $moduleRoot.FindName('dgHistory').ItemsSource = $history
            $moduleRoot.FindName('btnRepopulate').Visibility = 'Visible'
            $moduleRoot.FindName('txtResultBadge').Text = 'Pass'
            $moduleRoot.FindName('txtResultDetails').Text = 'MATCH (IsEquals): ''1.2.3'' vs ''1.2.3'' - a longer details text that wraps onto a second line when the window is narrow, to test the layout of the result panel.'
            $moduleRoot.FindName('rkvKeyPath').Text = 'SOFTWARE\Contoso\Application\Version'
            $moduleRoot.FindName('rkvValueName').Text = 'DisplayVersion'
            $moduleRoot.FindName('rkvExpected').Text = '1.2.3'
            $moduleRoot.FindName('rkKeyPath').Text = 'SOFTWARE\Contoso\Application'
            $moduleRoot.FindName('fileFilePath').Text = 'C:\Program Files\Contoso\Application'
            $moduleRoot.FindName('fileFileName').Text = 'application.exe'
            $moduleRoot.FindName('fileExpected').Text = '1.2.3.4'
            $moduleRoot.FindName('scriptText').Text = "if (Test-Path 'C:\Program Files\Contoso\Application\application.exe') { 'Installed' }"
            $moduleRoot.FindName('compoundConnector').Text = 'And'
            $list = $moduleRoot.FindName('compoundClauses')
            $list.Items.Clear()
            foreach ($line in $Config.SampleCompound) { [void]$list.Items.Add($line) }
        }
        else {
            $apps = New-Object 'System.Collections.ObjectModel.ObservableCollection[psobject]'
            foreach ($a in $Config.SampleApps) { $apps.Add([PSCustomObject]$a) }
            $grid = $moduleRoot.FindName('dgApps')
            $grid.ItemsSource = $apps
            $moduleRoot.FindName('txtFilter').Text = 'contoso'
            $grid.SelectedIndex = 0
            $a = $apps[0]
            $moduleRoot.FindName('btnUseForDetection').Visibility = 'Visible'
            $moduleRoot.FindName('btnCopyDetails').Visibility = 'Visible'
            $moduleRoot.FindName('dtlDisplayName').Text = $a.DisplayName
            $moduleRoot.FindName('dtlPublisher').Text = $a.Publisher
            $moduleRoot.FindName('dtlDisplayVersion').Text = $a.DisplayVersion
            $moduleRoot.FindName('dtlArchScope').Text = '{0} / {1}' -f $a.Architecture, $a.Scope
            $moduleRoot.FindName('dtlRegistryKey').Text = 'HKLM\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\{307F4E40-A4A5-4E15-B945-86FA88DC16F5}'
            $moduleRoot.FindName('dtlUninstallString').Text = 'MsiExec.exe /X{307F4E40-A4A5-4E15-B945-86FA88DC16F5}'
            $moduleRoot.FindName('dtlQuietUninstall').Text = 'MsiExec.exe /X{307F4E40-A4A5-4E15-B945-86FA88DC16F5} /qn'
            $moduleRoot.FindName('dtlInstallLocation').Text = 'C:\Program Files (x86)\Contoso\Application'
            $moduleRoot.FindName('dtlInstallDate').Text = '20261106'
        }
        Wait-Layout $Window
    }

    # Several layout passes with short pauses, so templates and animations of a newly shown
    # or newly enabled control have run before it is measured.
    function Wait-Layout {
        param($Window, [int]$Passes = 4)
        for ($i = 0; $i -lt $Passes; $i++) { Sync-Layout $Window; Start-Sleep -Milliseconds 40 }
        Sync-Layout $Window
    }

    function Sync-Layout {
        param($Window)
        $Window.UpdateLayout()
        $Window.Dispatcher.Invoke([action] {}, [System.Windows.Threading.DispatcherPriority]::ApplicationIdle)
    }

    function Get-VisualDescendants {
        param([System.Windows.DependencyObject]$Root)
        $stack = New-Object System.Collections.Generic.Stack[System.Windows.DependencyObject]
        $stack.Push($Root)
        while ($stack.Count -gt 0) {
            $n = $stack.Pop()
            $n
            $count = [System.Windows.Media.VisualTreeHelper]::GetChildrenCount($n)
            for ($i = $count - 1; $i -ge 0; $i--) { $stack.Push([System.Windows.Media.VisualTreeHelper]::GetChild($n, $i)) }
        }
    }

    function Test-VisibleInTree {
        param([System.Windows.UIElement]$Element)
        $d = $Element -as [System.Windows.DependencyObject]
        while ($d) {
            if ($d -is [System.Windows.UIElement] -and $d.Visibility -ne 'Visible') { return $false }
            $d = [System.Windows.Media.VisualTreeHelper]::GetParent($d)
        }
        return $true
    }
}

# Measures the main window with every configured view, and the Options dialog with its two panes, in one theme and at one
# size. Runs on the STA thread (Invoke-WpfSta). Contrast is read only at 1440x900; clipped text and off-screen controls
# are read at both sizes.
$script:MeasureAppScript = {
            $cfg = $Arguments.Config
            $w = Open-TestWindow -XamlPath (Join-Path $Root $cfg.Xaml) -Dark $Arguments.Dark -Width $Arguments.W -Height $Arguments.H
            # Contrast is read only at 1440x900. Clipped text and off-screen controls are read at both sizes.
            $full = $Arguments.W -ge 1440
            $out = @()
            Initialize-ShellRuntime -Window $w -Parts $Arguments.Parts -LogPath (Join-Path $Arguments.TempDir 'measure.log')
            foreach ($view in $cfg.Views) {
                Show-AppView -Window $w -Config $cfg -View $view
                # The shell's own theme code: theme, log label colour, sidebar button colours and title bar brushes.
                & $global:ApplyTheme -IsDark $Arguments.Dark
                Wait-Layout $w
                $out += [PSCustomObject]@{
                    View        = $view.Name
                    Text        = if ($full) { @(Get-TextContrast $w) } else { @() }
                    Clipped     = @(Get-ClippedText $w)
                    Offscreen   = @(Get-OffscreenControls $w)
                    Interactive = if ($full) { @(Get-InteractiveContrast $w) } else { @() }
                    TitleBar    = if ($full) { Get-TitleBarContrast $w } else { $null }
                }
            }
            $w.Close()

            # The Options dialog at its default size (1440x900 pass) and at its minimum size (1024x720 pass).
            $dw = if ($full) { 640 } else { 560 }
            $dh = if ($full) { 440 } else { 380 }
            $d = Open-TestWindow -XamlText $Arguments.DialogXaml -Dark $Arguments.Dark -Width $dw -Height $dh
            $d.Resources['MahApps.Brushes.WindowTitle.NonActive'] = $global:TitleBarInactiveBrush
            Set-DialogTheme -Dialog $d -IsDark $Arguments.Dark
            $d.FindName('txtAboutVersion').Text = 'v2026.11.06.0013'
            $d.FindName('txtAboutModule').Text = 'DetectionTesterCommon 2026.11.06.0013'
            $d.FindName('txtLogPath').Text = 'C:\Users\someone\AppData\Local\DetectionTester\logs\DetectionTester-20261106-101500.log'
            $d.FindName('txtLogDir').Text = 'C:\Users\someone\AppData\Local\DetectionTester\logs'
            foreach ($pane in @(@{ Name = 'Options dialog: About'; Show = 'paneAbout'; Hide = 'paneLogging' }, @{ Name = 'Options dialog: Logging'; Show = 'paneLogging'; Hide = 'paneAbout' })) {
                $d.FindName($pane.Show).Visibility = 'Visible'
                $d.FindName($pane.Hide).Visibility = 'Collapsed'
                Wait-Layout $d
                $out += [PSCustomObject]@{
                    View        = $pane.Name
                    Text        = if ($full) { @(Get-TextContrast $d) } else { @() }
                    Clipped     = @(Get-ClippedText $d)
                    Offscreen   = @(Get-OffscreenControls $d)
                    Interactive = if ($full) { @(Get-InteractiveContrast $d) } else { @() }
                    TitleBar    = if ($full) { Get-TitleBarContrast $d } else { $null }
                }
            }
            $d.Close()
            , $out
}

function Get-WpfSession {
    if ($script:HarnessSession -and $script:HarnessSession.RunspaceStateInfo.State -eq 'Opened') { return $script:HarnessSession }
    $rs = [runspacefactory]::CreateRunspace()
    $rs.ApartmentState = 'STA'
    $rs.ThreadOptions = 'ReuseThread'
    $rs.Open()
    $rs.SessionStateProxy.SetVariable('Root', $script:HarnessRoot)
    $ps = [powershell]::Create()
    $ps.Runspace = $rs
    [void]$ps.AddScript($script:HarnessBootstrap.ToString(), $false)
    [void]$ps.AddScript($script:MeasureFunctions.ToString(), $false)
    try {
        [void]$ps.Invoke()
        if ($ps.HadErrors) { throw (($ps.Streams.Error | ForEach-Object { $_.ToString() }) -join "`n") }
    }
    catch {
        $ps.Dispose()
        $rs.Dispose()
        throw
    }
    $ps.Dispose()
    $script:HarnessSession = $rs
    return $rs
}

function Close-WpfSession {
    if ($script:HarnessSession) { $script:HarnessSession.Dispose(); $script:HarnessSession = $null }
}

function Invoke-WpfSta {
    param(
        [Parameter(Mandatory)][scriptblock]$Script,
        [hashtable]$Arguments = @{},
        [ValidateRange(1, 3600)][int]$TimeoutSeconds = 300
    )
    $rs = Get-WpfSession
    $rs.SessionStateProxy.SetVariable('Arguments', $Arguments)
    $ps = [powershell]::Create()
    try {
        $ps.Runspace = $rs
        [void]$ps.AddScript($Script.ToString(), $false)
        # A stalled theme manager can block the script block for ever. Bound the wait, stop it, and drop the runspace.
        $async = $ps.BeginInvoke()
        if (-not $async.AsyncWaitHandle.WaitOne([TimeSpan]::FromSeconds($TimeoutSeconds))) {
            try { [void]$ps.BeginStop($null, $null).AsyncWaitHandle.WaitOne([TimeSpan]::FromSeconds(10)) } catch { Write-Verbose "Stop failed: $_" }
            Close-WpfSession
            throw "STA script block did not finish within $TimeoutSeconds seconds."
        }
        $out = @($ps.EndInvoke($async))
        if ($ps.HadErrors -and $ps.Streams.Error.Count -gt 0) {
            throw (($ps.Streams.Error | ForEach-Object { $_.ToString() }) -join "`n")
        }
        return $out
    }
    finally {
        $ps.Dispose()
    }
}
