# Lab-side UI Automation helpers. Dot-sourced by the runner inside the interactive session.
[Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseApprovedVerbs', '', Justification = 'Short helper names for the capture steps; the file is never imported as a module.')]
param()
Add-Type -AssemblyName UIAutomationClient, UIAutomationTypes, System.Drawing, System.Windows.Forms
Add-Type -TypeDefinition @'
using System; using System.Runtime.InteropServices;
public class NativeUi {
  [DllImport("user32.dll")] public static extern bool SetProcessDPIAware();
  [DllImport("user32.dll")] public static extern bool SetCursorPos(int x, int y);
  [DllImport("user32.dll")] public static extern void mouse_event(uint f, int dx, int dy, uint d, UIntPtr e);
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
function Maximize-Win($Win) {
    # The design size is the work area of a maximized window on a 1080p screen (D-20): the window fills it, nothing is hardcoded.
    $Win.GetCurrentPattern([System.Windows.Automation.WindowPattern]::Pattern).SetWindowVisualState([System.Windows.Automation.WindowVisualState]::Maximized)
    Start-Sleep -Milliseconds 1500
    $wa = [System.Windows.Forms.SystemInformation]::WorkingArea
    Log ('work area {0}x{1}' -f $wa.Width, $wa.Height)
}
function Shot($Win, [string]$Name) {
    # The mouse leaves the window, so that no button shows its hover state in the image.
    [void][NativeUi]::SetCursorPos(1900, 1060)
    Start-Sleep -Milliseconds 1300
    # A tooltip of the last button stays for a few seconds: wait until none shows. When one stays, touch a neutral point of the title bar and leave again.
    $tipCond = New-Object System.Windows.Automation.PropertyCondition($AE::ControlTypeProperty, [System.Windows.Automation.ControlType]::ToolTip)
    $end = (Get-Date).AddSeconds(8)
    while ((Get-Date) -lt $end -and $root.FindFirst('Descendants', $tipCond)) { Start-Sleep -Milliseconds 400 }
    if ($root.FindFirst('Descendants', $tipCond)) { [void][NativeUi]::SetCursorPos(700, 16); Start-Sleep -Milliseconds 400; [void][NativeUi]::SetCursorPos(1900, 1060); Start-Sleep -Seconds 2 }
    $r = $Win.Current.BoundingRectangle
    $x = [int][Math]::Max(0, $r.X); $y = [int][Math]::Max(0, $r.Y)
    $w = [int][Math]::Min($r.Width, [System.Windows.Forms.SystemInformation]::VirtualScreen.Width - $x)
    $h = [int][Math]::Min($r.Height, [System.Windows.Forms.SystemInformation]::VirtualScreen.Height - $y)
    $bmp = New-Object System.Drawing.Bitmap $w, $h; $g = [System.Drawing.Graphics]::FromImage($bmp)
    $g.CopyFromScreen($x, $y, 0, 0, (New-Object System.Drawing.Size $w, $h)); $g.Dispose()
    # The check of this shot (D-21): the mean luminance of a grid of pixels, and a warning for a blank image.
    $lum = 0.0; $cnt = 0; $same = 0; $first = $bmp.GetPixel(2, 2)
    for ($sy = 8; $sy -lt $h; $sy += 40) { for ($sx = 8; $sx -lt $w; $sx += 40) { $c = $bmp.GetPixel($sx, $sy); $lum += (0.2126 * $c.R + 0.7152 * $c.G + 0.0722 * $c.B) / 255; $cnt++; if ($c.ToArgb() -eq $first.ToArgb()) { $same++ } } }
    Log ('SAMPLE {0} meanLuminance {1:N2}' -f $Name, ($lum / [Math]::Max(1, $cnt)))
    if ($cnt -gt 0 -and ($same / $cnt) -gt 0.98) { Log ('WARN {0} looks blank' -f $Name) }
    $bmp.Save((Join-Path $OutDir ($Name + '.png')), [System.Drawing.Imaging.ImageFormat]::Png); $bmp.Dispose()
    Log ('shot {0} {1}x{2}' -f $Name, $w, $h)
}
# The capture stops only the processes that it started, by process id (vault: testing-policy, "Process rules on a lab VM").
# Each run records its processes in a ledger in its own run folder. No process is ever stopped by a command-line pattern.
$script:RunLedger = Join-Path $PSScriptRoot 'run.pids'
$script:AppProcs = @{}
function Register-RunProcess($Process, [string]$Key) {
    $script:AppProcs[$Key] = $Process
    ('{0}|{1}' -f $Process.Id, $Process.StartTime.Ticks) | Add-Content -LiteralPath $script:RunLedger
}
function Start-App([string]$Script, [string]$Title) {
    [void][NativeUi]::SetCursorPos(1900, 1060)
    $script:proc = Start-Process -FilePath 'powershell.exe' -ArgumentList @('-NoProfile', '-ExecutionPolicy', 'Bypass', '-STA', '-File', (Join-Path $appDir $Script)) -PassThru
    Register-RunProcess $script:proc $Script
    $null = Find-Top $Title 60; Start-Sleep -Seconds 3; Find-Top $Title 10
}
function Stop-App([string]$Script) {
    # Stop only the process that Start-App started for this script, by its process id.
    $p = $script:AppProcs[$Script]
    if ($p -and -not $p.HasExited) { Stop-Process -Id $p.Id -Force -ErrorAction SilentlyContinue }
}
function Section([string]$Name, [scriptblock]$Body) {
    # A run with -Section lists the sections to run in $Capture.Sections.
    $selected = @($Capture.Sections | Where-Object { $_ })
    if ($selected.Count -and $Name -notin $selected) { Log ('== ' + $Name + ' not selected'); return }
    Log ('== ' + $Name)
    try { & $Body } catch { Log ('ERROR in ' + $Name + ': ' + $_.Exception.Message + ' @ ' + $_.InvocationInfo.ScriptLineNumber) }
}
function Set-Text($El, [string]$Text) { $El.GetCurrentPattern([System.Windows.Automation.ValuePattern]::Pattern).SetValue($Text) }
function Expand-El($El) { $El.GetCurrentPattern([System.Windows.Automation.ExpandCollapsePattern]::Pattern).Expand() }
function Collapse-El($El) { $El.GetCurrentPattern([System.Windows.Automation.ExpandCollapsePattern]::Pattern).Collapse() }
function Get-ElText($El) { try { $El.GetCurrentPattern([System.Windows.Automation.ValuePattern]::Pattern).Current.Value } catch { $El.Current.Name } }
function Click-At([int]$x, [int]$y) { [void][NativeUi]::SetCursorPos($x, $y); Start-Sleep -Milliseconds 150; [NativeUi]::mouse_event(2, 0, 0, 0, [UIntPtr]::Zero); [NativeUi]::mouse_event(4, 0, 0, 0, [UIntPtr]::Zero); Start-Sleep -Milliseconds 250 }
function Click-El($El) { $r = $El.Current.BoundingRectangle; Click-At ([int]($r.X + $r.Width / 2)) ([int]($r.Y + $r.Height / 2)) }
function Send-Keys([string]$k) { [System.Windows.Forms.SendKeys]::SendWait($k); Start-Sleep -Milliseconds 400 }
function Focus-Win($Win) { try { [void][NativeUi]::SetForegroundWindow([IntPtr]$Win.Current.NativeWindowHandle) } catch { $null = $_ }; Start-Sleep -Milliseconds 300 }
function Close-Win($Win) { $Win.GetCurrentPattern([System.Windows.Automation.WindowPattern]::Pattern).Close() }
function Find-All([string]$Name) {
    $cond = New-Object System.Windows.Automation.AndCondition((New-Object System.Windows.Automation.PropertyCondition($AE::NameProperty, $Name)), (New-Object System.Windows.Automation.PropertyCondition($AE::ControlTypeProperty, [System.Windows.Automation.ControlType]::Window)))
    @($root.FindAll('Descendants', $cond))
}
function File-Dialog([string]$Title, [string]$Path) {
    # Wait for the file dialog. Stop the capture when it does not show, so that a file that was never chosen cannot count as a success.
    $d = $null; $end = (Get-Date).AddSeconds(20)
    while (-not $d -and (Get-Date) -lt $end) {
        $d = @(Find-All $Title) | Select-Object -First 1
        if (-not $d) { $d = @($root.FindAll('Children', [System.Windows.Automation.Condition]::TrueCondition)) | Where-Object { $_.Current.ClassName -eq '#32770' } | Select-Object -First 1 }
        if (-not $d) { Start-Sleep -Milliseconds 400 }
    }
    if (-not $d) { Shot-Screen ('fail-dialog-' + ($Title -replace '\W', '_')); throw ('The file dialog did not show: ' + $Title) }
    Start-Sleep -Milliseconds 800
    Focus-Win $d
    Send-Keys ($Path -replace '([+^%~{}\[\]()])', '{$1}')
    Send-Keys '{ENTER}'; Start-Sleep -Seconds 2
}
function Dump-Tree($Scope, [string]$Name) {
    $out = foreach ($e in $Scope.FindAll('Descendants', [System.Windows.Automation.Condition]::TrueCondition)) {
        $c = $e.Current; if ($c.Name -or $c.AutomationId) { '{0}|{1}|{2}|{3}' -f $c.ControlType.ProgrammaticName.Replace('ControlType.', ''), $c.AutomationId, $c.Name, $c.ClassName }
    }
    $out | Set-Content (Join-Path $OutDir ($Name + '.tree.txt'))
}
