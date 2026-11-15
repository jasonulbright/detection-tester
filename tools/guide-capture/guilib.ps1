# Lab-side UI Automation helpers. Dot-sourced by the runner inside the interactive session.
[Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseApprovedVerbs', '', Justification = 'Short helper names for the capture steps; the file is never imported as a module.')]
param()
Add-Type -AssemblyName UIAutomationClient, UIAutomationTypes, System.Drawing, System.Windows.Forms
Add-Type -TypeDefinition @'
using System; using System.Runtime.InteropServices;
public class NativeUi {
  [DllImport("user32.dll")] public static extern bool SetProcessDPIAware();
  [DllImport("user32.dll")] public static extern bool SetCursorPos(int x, int y);
  [DllImport("user32.dll")] public static extern bool SetForegroundWindow(IntPtr h);
  [DllImport("user32.dll")] public static extern bool MoveWindow(IntPtr h, int x, int y, int w, int ht, bool r);
}
'@
[void][NativeUi]::SetProcessDPIAware()
$AE = [System.Windows.Automation.AutomationElement]
$root = $AE::RootElement
function Log([string]$m) { ('{0}  {1}' -f (Get-Date -Format 'HH:mm:ss'), $m) | Add-Content -LiteralPath (Join-Path $OutDir 'run.log') }
function Shot-Screen([string]$Name) {
    $s = [System.Windows.Forms.SystemInformation]::VirtualScreen
    $b = New-Object System.Drawing.Bitmap $s.Width, $s.Height; $g = [System.Drawing.Graphics]::FromImage($b); $g.CopyFromScreen(0, 0, 0, 0, $b.Size); $g.Dispose()
    $b.Save((Join-Path $OutDir ($Name + '.png'))); $b.Dispose(); Log ('screen ' + $Name)
}
function Find-Top([string]$Name, [int]$Seconds = 30) {
    $cond = New-Object System.Windows.Automation.AndCondition((New-Object System.Windows.Automation.PropertyCondition($AE::NameProperty, $Name)), (New-Object System.Windows.Automation.PropertyCondition($AE::ControlTypeProperty, [System.Windows.Automation.ControlType]::Window)))
    $end = (Get-Date).AddSeconds($Seconds)
    do { $el = $root.FindFirst('Descendants', $cond); if ($el) { return $el }; Start-Sleep -Milliseconds 400 } while ((Get-Date) -lt $end)
    $names = @($root.FindAll('Children', [System.Windows.Automation.Condition]::TrueCondition) | ForEach-Object { $_.Current.Name + '|' + $_.Current.ClassName }) -join ' ; '
    Log ('top windows: ' + $names); Shot-Screen ('fail-' + ($Name -replace '\W', '_')); throw "window not found: $Name"
}
function Get-El($Scope, [string]$Id, [switch]$Quiet) {
    $el = $Scope.FindFirst('Descendants', (New-Object System.Windows.Automation.PropertyCondition($AE::AutomationIdProperty, $Id)))
    if (-not $el) { $el = $Scope.FindFirst('Descendants', (New-Object System.Windows.Automation.PropertyCondition($AE::NameProperty, $Id))) }
    if (-not $el -and -not $Quiet) { throw "control not found: $Id" }
    $el
}
function Find-Role($Scope, [string]$Type, [string]$Name) {
    $ct = [System.Windows.Automation.ControlType]::($Type)
    $c = New-Object System.Windows.Automation.AndCondition((New-Object System.Windows.Automation.PropertyCondition($AE::ControlTypeProperty, $ct)), (New-Object System.Windows.Automation.PropertyCondition($AE::NameProperty, $Name)))
    $Scope.FindFirst('Descendants', $c)
}
function Invoke-El($El) { [void][NativeUi]::SetCursorPos(1900, 1060); $El.GetCurrentPattern([System.Windows.Automation.InvokePattern]::Pattern).Invoke() }
function Toggle-El($El) { $El.GetCurrentPattern([System.Windows.Automation.TogglePattern]::Pattern).Toggle() }
function Select-El($El) { $El.GetCurrentPattern([System.Windows.Automation.SelectionItemPattern]::Pattern).Select() }
function Wait-Until([scriptblock]$Test, [int]$Seconds = 120) { $end = (Get-Date).AddSeconds($Seconds); while ((Get-Date) -lt $end) { if (& $Test) { return $true }; Start-Sleep -Milliseconds 500 }; $false }
function Size-Win($Win, [int]$W = 1440, [int]$H = 900) {
    $hh = [IntPtr]$Win.Current.NativeWindowHandle
    $sw = [System.Windows.Forms.SystemInformation]::VirtualScreen
    $w2 = [Math]::Min($W, $sw.Width); $h2 = [Math]::Min($H, $sw.Height - 40)
    [void][NativeUi]::MoveWindow($hh, 0, 0, $w2, $h2, $true); Start-Sleep -Milliseconds 900
}
function Shot($Win, [string]$Name) {
    # The mouse leaves the window, so that no button shows its hover state in the image.
    [void][NativeUi]::SetCursorPos(1900, 1060)
    Start-Sleep -Milliseconds 1300
    $r = $Win.Current.BoundingRectangle
    $x = [int][Math]::Max(0, $r.X); $y = [int][Math]::Max(0, $r.Y)
    $w = [int][Math]::Min($r.Width, [System.Windows.Forms.SystemInformation]::VirtualScreen.Width - $x)
    $h = [int][Math]::Min($r.Height, [System.Windows.Forms.SystemInformation]::VirtualScreen.Height - $y)
    $bmp = New-Object System.Drawing.Bitmap $w, $h; $g = [System.Drawing.Graphics]::FromImage($bmp)
    $g.CopyFromScreen($x, $y, 0, 0, (New-Object System.Drawing.Size $w, $h)); $g.Dispose()
    $bmp.Save((Join-Path $OutDir ($Name + '.png')), [System.Drawing.Imaging.ImageFormat]::Png); $bmp.Dispose()
    Log ('shot {0} {1}x{2}' -f $Name, $w, $h)
}
function Start-App([string]$Script, [string]$Title) {
    [void][NativeUi]::SetCursorPos(1900, 1060)
    $script:proc = Start-Process -FilePath 'powershell.exe' -ArgumentList @('-NoProfile', '-ExecutionPolicy', 'Bypass', '-STA', '-File', (Join-Path $appDir $Script)) -PassThru
    $null = Find-Top $Title 60; Start-Sleep -Seconds 3; Find-Top $Title 10
}
function Stop-App([string]$Script) {
    Get-CimInstance Win32_Process -Filter "Name='powershell.exe'" | Where-Object { $_.CommandLine -like "*$Script*" } | ForEach-Object { Stop-Process -Id $_.ProcessId -Force -ErrorAction SilentlyContinue }
}
function Section([string]$Name, [scriptblock]$Body) {
    # A run with -Section lists the sections to run in $Capture.Sections.
    $selected = @($Capture.Sections | Where-Object { $_ })
    if ($selected.Count -and $Name -notin $selected) { Log ('== ' + $Name + ' not selected'); return }
    Log ('== ' + $Name)
    try { & $Body } catch { Log ('ERROR in ' + $Name + ': ' + $_.Exception.Message + ' @ ' + $_.InvocationInfo.ScriptLineNumber) }
}
