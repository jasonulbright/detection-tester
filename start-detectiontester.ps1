<#
.SYNOPSIS
    Detection Method Testing Tool for MECM application detection rules.

.DESCRIPTION
    WinForms PowerShell GUI for testing MECM detection methods against the local machine
    without deploying through MECM. Supports RegistryKeyValue, RegistryKey, File, Script,
    and Compound detection types. Includes an Installed Applications browser for quick
    detection rule creation.

.EXAMPLE
    .\start-detectiontester.ps1

.NOTES
    Requirements:
      - PowerShell 5.1
      - .NET Framework 4.8+
      - Windows Forms (System.Windows.Forms)

    No admin rights and no MECM connection required.

    ScriptName : start-detectiontester.ps1
    Purpose    : Test MECM detection methods locally
    Version    : 1.0.2
    Updated    : 2026-03-17
#>

param()

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

[System.Windows.Forms.Application]::EnableVisualStyles()
try { [System.Windows.Forms.Application]::SetCompatibleTextRenderingDefault($false) } catch { }

$moduleRoot = Join-Path $PSScriptRoot "Module"
Import-Module (Join-Path $moduleRoot "DetectionTesterCommon.psd1") -Force -DisableNameChecking

# Initialize logging
$toolLogFolder = Join-Path $PSScriptRoot "Logs"
if (-not (Test-Path -LiteralPath $toolLogFolder)) {
    New-Item -ItemType Directory -Path $toolLogFolder -Force | Out-Null
}
$toolLogPath = Join-Path $toolLogFolder ("DetectionTester-{0}.log" -f (Get-Date -Format 'yyyyMMdd-HHmmss'))
Initialize-Logging -LogPath $toolLogPath

# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

function Set-ModernButtonStyle {
    param(
        [Parameter(Mandatory)][System.Windows.Forms.Button]$Button,
        [Parameter(Mandatory)][System.Drawing.Color]$BackColor
    )
    $Button.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
    $Button.FlatAppearance.BorderSize = 0
    $Button.BackColor = $BackColor
    $Button.ForeColor = [System.Drawing.Color]::White
    $Button.UseVisualStyleBackColor = $false
    $Button.Cursor = [System.Windows.Forms.Cursors]::Hand
    $hover = [System.Drawing.Color]::FromArgb([Math]::Max(0, $BackColor.R - 18), [Math]::Max(0, $BackColor.G - 18), [Math]::Max(0, $BackColor.B - 18))
    $down  = [System.Drawing.Color]::FromArgb([Math]::Max(0, $BackColor.R - 36), [Math]::Max(0, $BackColor.G - 36), [Math]::Max(0, $BackColor.B - 36))
    $Button.FlatAppearance.MouseOverBackColor = $hover
    $Button.FlatAppearance.MouseDownBackColor = $down
}

function Enable-DoubleBuffer {
    param([Parameter(Mandatory)][System.Windows.Forms.Control]$Control)
    $prop = $Control.GetType().GetProperty("DoubleBuffered", [System.Reflection.BindingFlags] "Instance,NonPublic")
    if ($prop) { $prop.SetValue($Control, $true, $null) | Out-Null }
}

function New-ThemedGrid {
    param([switch]$MultiSelect)

    $g = New-Object System.Windows.Forms.DataGridView
    $g.Dock = [System.Windows.Forms.DockStyle]::Fill
    $g.ReadOnly = $true
    $g.AllowUserToAddRows = $false
    $g.AllowUserToDeleteRows = $false
    $g.AllowUserToResizeRows = $false
    $g.SelectionMode = [System.Windows.Forms.DataGridViewSelectionMode]::FullRowSelect
    $g.MultiSelect = [bool]$MultiSelect
    $g.AutoGenerateColumns = $false
    $g.RowHeadersVisible = $false
    $g.BackgroundColor = $clrPanelBg
    $g.BorderStyle = [System.Windows.Forms.BorderStyle]::None
    $g.CellBorderStyle = [System.Windows.Forms.DataGridViewCellBorderStyle]::SingleHorizontal
    $g.GridColor = $clrGridLine
    $g.ColumnHeadersDefaultCellStyle.BackColor = $clrAccent
    $g.ColumnHeadersDefaultCellStyle.ForeColor = [System.Drawing.Color]::White
    $g.ColumnHeadersDefaultCellStyle.Font = New-Object System.Drawing.Font("Segoe UI", 8, [System.Drawing.FontStyle]::Bold)
    $g.ColumnHeadersDefaultCellStyle.Padding = New-Object System.Windows.Forms.Padding(4)
    $g.ColumnHeadersDefaultCellStyle.WrapMode = [System.Windows.Forms.DataGridViewTriState]::False
    $g.ColumnHeadersHeightSizeMode = [System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode]::DisableResizing
    $g.ColumnHeadersHeight = 32
    $g.ColumnHeadersBorderStyle = [System.Windows.Forms.DataGridViewHeaderBorderStyle]::Single
    $g.EnableHeadersVisualStyles = $false
    $g.DefaultCellStyle.Font = New-Object System.Drawing.Font("Segoe UI", 9)
    $g.DefaultCellStyle.ForeColor = $clrGridText
    $g.DefaultCellStyle.BackColor = $clrPanelBg
    $g.DefaultCellStyle.Padding = New-Object System.Windows.Forms.Padding(2)
    $selBg = if ($script:Prefs.DarkMode) { [System.Drawing.Color]::FromArgb(38, 79, 120) } else { [System.Drawing.Color]::FromArgb(0, 120, 215) }
    $g.DefaultCellStyle.SelectionBackColor = $selBg
    $g.DefaultCellStyle.SelectionForeColor = [System.Drawing.Color]::White
    $g.RowTemplate.Height = 26
    $g.AlternatingRowsDefaultCellStyle.BackColor = $clrGridAlt
    Enable-DoubleBuffer -Control $g
    return $g
}

function Save-WindowState {
    $statePath = Join-Path $PSScriptRoot "DetectionTester.windowstate.json"
    $state = @{
        X = $form.Location.X; Y = $form.Location.Y
        Width = $form.Size.Width; Height = $form.Size.Height
        Maximized = ($form.WindowState -eq [System.Windows.Forms.FormWindowState]::Maximized)
        ActiveTab = $tabMain.SelectedIndex
    }
    $state | ConvertTo-Json | Set-Content -LiteralPath $statePath -Encoding UTF8
}

function Restore-WindowState {
    $statePath = Join-Path $PSScriptRoot "DetectionTester.windowstate.json"
    if (-not (Test-Path -LiteralPath $statePath)) { return }
    try {
        $state = Get-Content -LiteralPath $statePath -Raw | ConvertFrom-Json
        if ($state.Maximized) { $form.WindowState = [System.Windows.Forms.FormWindowState]::Maximized }
        else {
            $screen = [System.Windows.Forms.Screen]::FromPoint((New-Object System.Drawing.Point($state.X, $state.Y)))
            $bounds = $screen.WorkingArea
            $x = [Math]::Max($bounds.X, [Math]::Min($state.X, $bounds.Right - 200))
            $y = [Math]::Max($bounds.Y, [Math]::Min($state.Y, $bounds.Bottom - 100))
            $form.Location = New-Object System.Drawing.Point($x, $y)
            $form.Size = New-Object System.Drawing.Size([Math]::Max($form.MinimumSize.Width, $state.Width), [Math]::Max($form.MinimumSize.Height, $state.Height))
        }
        if ($null -ne $state.ActiveTab -and $state.ActiveTab -ge 0 -and $state.ActiveTab -lt $tabMain.TabCount) {
            $tabMain.SelectedIndex = [int]$state.ActiveTab
        }
    } catch { }
}

# ---------------------------------------------------------------------------
# Preferences
# ---------------------------------------------------------------------------

function Get-DetectionTesterPreferences {
    $prefsPath = Join-Path $PSScriptRoot "DetectionTester.prefs.json"
    $defaults = @{ DarkMode = $false }
    if (Test-Path -LiteralPath $prefsPath) {
        try {
            $loaded = Get-Content -LiteralPath $prefsPath -Raw | ConvertFrom-Json
            if ($null -ne $loaded.DarkMode) { $defaults.DarkMode = [bool]$loaded.DarkMode }
        } catch { }
    }
    return $defaults
}

function Save-DetectionTesterPreferences {
    param([hashtable]$Prefs)
    $prefsPath = Join-Path $PSScriptRoot "DetectionTester.prefs.json"
    $Prefs | ConvertTo-Json | Set-Content -LiteralPath $prefsPath -Encoding UTF8
}

$script:Prefs = Get-DetectionTesterPreferences

# ---------------------------------------------------------------------------
# Colors (theme-aware)
# ---------------------------------------------------------------------------

$clrAccent = [System.Drawing.Color]::FromArgb(0, 120, 212)

if ($script:Prefs.DarkMode) {
    $clrFormBg   = [System.Drawing.Color]::FromArgb(30, 30, 30)
    $clrPanelBg  = [System.Drawing.Color]::FromArgb(40, 40, 40)
    $clrHint     = [System.Drawing.Color]::FromArgb(140, 140, 140)
    $clrGridAlt  = [System.Drawing.Color]::FromArgb(48, 48, 48)
    $clrGridLine = [System.Drawing.Color]::FromArgb(60, 60, 60)
    $clrDetailBg = [System.Drawing.Color]::FromArgb(45, 45, 45)
    $clrSepLine  = [System.Drawing.Color]::FromArgb(55, 55, 55)
    $clrText     = [System.Drawing.Color]::FromArgb(220, 220, 220)
    $clrGridText = [System.Drawing.Color]::FromArgb(220, 220, 220)
    $clrErrText  = [System.Drawing.Color]::FromArgb(255, 100, 100)
    $clrOkText   = [System.Drawing.Color]::FromArgb(80, 200, 80)
} else {
    $clrFormBg   = [System.Drawing.Color]::FromArgb(245, 246, 248)
    $clrPanelBg  = [System.Drawing.Color]::White
    $clrHint     = [System.Drawing.Color]::FromArgb(140, 140, 140)
    $clrGridAlt  = [System.Drawing.Color]::FromArgb(248, 250, 252)
    $clrGridLine = [System.Drawing.Color]::FromArgb(230, 230, 230)
    $clrDetailBg = [System.Drawing.Color]::FromArgb(250, 250, 250)
    $clrSepLine  = [System.Drawing.Color]::FromArgb(218, 220, 224)
    $clrText     = [System.Drawing.Color]::Black
    $clrGridText = [System.Drawing.Color]::Black
    $clrErrText  = [System.Drawing.Color]::FromArgb(180, 0, 0)
    $clrOkText   = [System.Drawing.Color]::FromArgb(34, 139, 34)
}

# Dark mode ToolStrip renderer
if ($script:Prefs.DarkMode) {
    if (-not ('DarkToolStripRenderer' -as [type])) {
        $rendererCs = @(
            'using System.Drawing;', 'using System.Windows.Forms;',
            'public class DarkToolStripRenderer : ToolStripProfessionalRenderer {',
            '    private Color _bg;',
            '    public DarkToolStripRenderer(Color bg) : base() { _bg = bg; }',
            '    protected override void OnRenderToolStripBorder(ToolStripRenderEventArgs e) { }',
            '    protected override void OnRenderToolStripBackground(ToolStripRenderEventArgs e) { using (var b = new SolidBrush(_bg)) { e.Graphics.FillRectangle(b, e.AffectedBounds); } }',
            '    protected override void OnRenderMenuItemBackground(ToolStripItemRenderEventArgs e) { if (e.Item.Selected || e.Item.Pressed) { using (var b = new SolidBrush(Color.FromArgb(60, 60, 60))) { e.Graphics.FillRectangle(b, new Rectangle(Point.Empty, e.Item.Size)); } } }',
            '    protected override void OnRenderSeparator(ToolStripSeparatorRenderEventArgs e) { int y = e.Item.Height / 2; using (var p = new Pen(Color.FromArgb(70, 70, 70))) { e.Graphics.DrawLine(p, 0, y, e.Item.Width, y); } }',
            '    protected override void OnRenderImageMargin(ToolStripRenderEventArgs e) { using (var b = new SolidBrush(_bg)) { e.Graphics.FillRectangle(b, e.AffectedBounds); } }',
            '}'
        ) -join "`r`n"
        Add-Type -ReferencedAssemblies System.Windows.Forms, System.Drawing -TypeDefinition $rendererCs
    }
    $script:DarkRenderer = New-Object DarkToolStripRenderer($clrPanelBg)
}

# ---------------------------------------------------------------------------
# Preferences dialog
# ---------------------------------------------------------------------------

function Show-PreferencesDialog {
    $scriptFile = Join-Path $PSScriptRoot "start-detectiontester.ps1"
    $dlg = New-Object System.Windows.Forms.Form
    $dlg.Text = "Preferences"; $dlg.Size = New-Object System.Drawing.Size(440, 200)
    $dlg.MinimumSize = $dlg.Size; $dlg.MaximumSize = $dlg.Size
    $dlg.StartPosition = "CenterParent"; $dlg.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedDialog
    $dlg.MaximizeBox = $false; $dlg.MinimizeBox = $false; $dlg.ShowInTaskbar = $false
    $dlg.Font = New-Object System.Drawing.Font("Segoe UI", 9.5); $dlg.BackColor = $clrFormBg

    $grpApp = New-Object System.Windows.Forms.GroupBox
    $grpApp.Text = "Appearance"; $grpApp.SetBounds(16, 12, 392, 60)
    $grpApp.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
    $grpApp.ForeColor = $clrText; $grpApp.BackColor = $clrFormBg
    if ($script:Prefs.DarkMode) { $grpApp.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat; $grpApp.ForeColor = $clrSepLine }
    $dlg.Controls.Add($grpApp)

    $chkDark = New-Object System.Windows.Forms.CheckBox
    $chkDark.Text = "Enable dark mode (requires restart)"; $chkDark.Font = New-Object System.Drawing.Font("Segoe UI", 9)
    $chkDark.AutoSize = $true; $chkDark.Location = New-Object System.Drawing.Point(14, 24)
    $chkDark.Checked = $script:Prefs.DarkMode; $chkDark.ForeColor = $clrText; $chkDark.BackColor = $clrFormBg
    if ($script:Prefs.DarkMode) { $chkDark.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat; $chkDark.ForeColor = [System.Drawing.Color]::FromArgb(170, 170, 170) }
    $grpApp.Controls.Add($chkDark)

    $btnSave = New-Object System.Windows.Forms.Button; $btnSave.Text = "Save"; $btnSave.SetBounds(220, 110, 90, 32)
    $btnSave.Font = New-Object System.Drawing.Font("Segoe UI", 9); Set-ModernButtonStyle -Button $btnSave -BackColor $clrAccent
    $dlg.Controls.Add($btnSave)
    $btnCancel = New-Object System.Windows.Forms.Button; $btnCancel.Text = "Cancel"; $btnCancel.SetBounds(318, 110, 90, 32)
    $btnCancel.Font = New-Object System.Drawing.Font("Segoe UI", 9); $btnCancel.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
    $btnCancel.ForeColor = $clrText; $btnCancel.BackColor = $clrFormBg
    $dlg.Controls.Add($btnCancel)

    $btnSave.Add_Click({
        $needsRestart = ($chkDark.Checked -ne $script:Prefs.DarkMode)
        $script:Prefs.DarkMode = $chkDark.Checked
        Save-DetectionTesterPreferences -Prefs $script:Prefs
        $dlg.DialogResult = [System.Windows.Forms.DialogResult]::OK; $dlg.Close()
        if ($needsRestart) {
            $result = [System.Windows.Forms.MessageBox]::Show("Theme change requires a restart. Restart now?", "Restart Required", [System.Windows.Forms.MessageBoxButtons]::YesNo, [System.Windows.Forms.MessageBoxIcon]::Question)
            if ($result -eq [System.Windows.Forms.DialogResult]::Yes) {
                Save-WindowState
                Start-Process powershell.exe -ArgumentList @('-ExecutionPolicy', 'Bypass', '-File', "`"$scriptFile`"")
                $form.Close()
            }
        }
    })
    $btnCancel.Add_Click({ $dlg.DialogResult = [System.Windows.Forms.DialogResult]::Cancel; $dlg.Close() })
    $dlg.AcceptButton = $btnSave; $dlg.CancelButton = $btnCancel
    $dlg.ShowDialog($form) | Out-Null; $dlg.Dispose()
}

# ---------------------------------------------------------------------------
# Form
# ---------------------------------------------------------------------------

$form = New-Object System.Windows.Forms.Form
$form.Text = "Detection Method Tester"; $form.Size = New-Object System.Drawing.Size(1100, 750)
$form.MinimumSize = New-Object System.Drawing.Size(900, 600); $form.StartPosition = "CenterScreen"
$form.Font = New-Object System.Drawing.Font("Segoe UI", 9.5); $form.BackColor = $clrFormBg

# ---------------------------------------------------------------------------
# StatusStrip (Dock:Bottom)
# ---------------------------------------------------------------------------

$statusStrip = New-Object System.Windows.Forms.StatusStrip
$statusStrip.BackColor = $clrPanelBg; $statusStrip.ForeColor = $clrText; $statusStrip.SizingGrip = $false
if ($script:Prefs.DarkMode -and $script:DarkRenderer) { $statusStrip.Renderer = $script:DarkRenderer }
$statusLabel = New-Object System.Windows.Forms.ToolStripStatusLabel; $statusLabel.Text = "Ready"
$statusLabel.Spring = $true; $statusLabel.TextAlign = [System.Drawing.ContentAlignment]::MiddleLeft
$statusStrip.Items.Add($statusLabel) | Out-Null
$form.Controls.Add($statusStrip)

# ---------------------------------------------------------------------------
# MenuStrip (Dock:Top)
# ---------------------------------------------------------------------------

$menuStrip = New-Object System.Windows.Forms.MenuStrip
$menuStrip.BackColor = $clrPanelBg; $menuStrip.ForeColor = $clrText
if ($script:Prefs.DarkMode -and $script:DarkRenderer) { $menuStrip.Renderer = $script:DarkRenderer }

$menuFile = New-Object System.Windows.Forms.ToolStripMenuItem("&File")
$menuFile.ForeColor = $clrText
$menuFilePrefs = New-Object System.Windows.Forms.ToolStripMenuItem("&Preferences...")
$menuFilePrefs.ForeColor = $clrText
$menuFilePrefs.Add_Click({ Show-PreferencesDialog })
$menuFile.DropDownItems.Add($menuFilePrefs) | Out-Null
$menuFile.DropDownItems.Add((New-Object System.Windows.Forms.ToolStripSeparator)) | Out-Null
$menuFileExit = New-Object System.Windows.Forms.ToolStripMenuItem("E&xit")
$menuFileExit.ForeColor = $clrText
$menuFileExit.Add_Click({ $form.Close() })
$menuFile.DropDownItems.Add($menuFileExit) | Out-Null
$menuStrip.Items.Add($menuFile) | Out-Null

$menuHelp = New-Object System.Windows.Forms.ToolStripMenuItem("&Help")
$menuHelp.ForeColor = $clrText
$menuHelpAbout = New-Object System.Windows.Forms.ToolStripMenuItem("&About Detection Method Tester...")
$menuHelpAbout.ForeColor = $clrText
$menuHelpAbout.Add_Click({
    $aboutText = @(
        "Detection Method Tester v1.0.2", "",
        "Test MECM application detection methods against the local machine",
        "without deploying through MECM.", "",
        "Supports: RegistryKeyValue, RegistryKey, File, Script, Compound", "",
        "Copyright (c) 2026 Jason Ulbright", "MIT License"
    ) -join "`r`n"
    [System.Windows.Forms.MessageBox]::Show($aboutText, "About Detection Method Tester", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information) | Out-Null
})
$menuHelp.DropDownItems.Add($menuHelpAbout) | Out-Null
$menuStrip.Items.Add($menuHelp) | Out-Null

$form.Controls.Add($menuStrip)
$form.MainMenuStrip = $menuStrip
$menuStrip.SendToBack()

# ---------------------------------------------------------------------------
# TabControl (Dock:Fill)
# ---------------------------------------------------------------------------

$tabMain = New-Object System.Windows.Forms.TabControl
$tabMain.Dock = [System.Windows.Forms.DockStyle]::Fill
$tabMain.Font = New-Object System.Drawing.Font("Segoe UI", 9.5, [System.Drawing.FontStyle]::Bold)
$tabMain.SizeMode = [System.Windows.Forms.TabSizeMode]::Fixed
$tabMain.ItemSize = New-Object System.Drawing.Size(160, 30)
$tabMain.DrawMode = [System.Windows.Forms.TabDrawMode]::OwnerDrawFixed

$tabMain.Add_DrawItem({
    param($s, $e)
    $e.Graphics.TextRenderingHint = [System.Drawing.Text.TextRenderingHint]::ClearTypeGridFit
    $tab = $s.TabPages[$e.Index]
    $sel = ($s.SelectedIndex -eq $e.Index)
    $bg = if ($script:Prefs.DarkMode) {
        if ($sel) { $clrAccent } else { $clrPanelBg }
    } else {
        if ($sel) { $clrAccent } else { [System.Drawing.Color]::FromArgb(240, 240, 240) }
    }
    $fg = if ($sel) { [System.Drawing.Color]::White } else { $clrText }
    $bb = New-Object System.Drawing.SolidBrush($bg)
    $e.Graphics.FillRectangle($bb, $e.Bounds)
    $ft = New-Object System.Drawing.Font("Segoe UI", 8, [System.Drawing.FontStyle]::Bold)
    $sf = New-Object System.Drawing.StringFormat
    $sf.Alignment = [System.Drawing.StringAlignment]::Near
    $sf.LineAlignment = [System.Drawing.StringAlignment]::Far
    $sf.FormatFlags = [System.Drawing.StringFormatFlags]::NoWrap
    $tr = New-Object System.Drawing.RectangleF(($e.Bounds.X + 8), $e.Bounds.Y, ($e.Bounds.Width - 12), ($e.Bounds.Height - 3))
    $tb = New-Object System.Drawing.SolidBrush($fg)
    $e.Graphics.DrawString($tab.Text, $ft, $tb, $tr, $sf)
    $bb.Dispose(); $tb.Dispose(); $ft.Dispose(); $sf.Dispose()
})

$form.Controls.Add($tabMain)
$tabMain.BringToFront()

# ===========================================================================
# TAB 1: Detection Tester
# ===========================================================================

$tabDetection = New-Object System.Windows.Forms.TabPage
$tabDetection.Text = "Detection Tester"
$tabDetection.BackColor = $clrFormBg
$tabMain.TabPages.Add($tabDetection)

# -- History DataTable (used by grid and export) --
$script:HistoryTable = New-Object System.Data.DataTable
[void]$script:HistoryTable.Columns.Add("Time", [string])
[void]$script:HistoryTable.Columns.Add("Type", [string])
[void]$script:HistoryTable.Columns.Add("Target", [string])
[void]$script:HistoryTable.Columns.Add("Expected", [string])
[void]$script:HistoryTable.Columns.Add("Found", [string])
[void]$script:HistoryTable.Columns.Add("Result", [string])

# -- Top panel: detection type selector --
$pnlDetType = New-Object System.Windows.Forms.Panel
$pnlDetType.Dock = [System.Windows.Forms.DockStyle]::Top
$pnlDetType.Height = 40; $pnlDetType.BackColor = $clrFormBg
$pnlDetType.Padding = New-Object System.Windows.Forms.Padding(12, 8, 12, 4)
$tabDetection.Controls.Add($pnlDetType)

$lblDetType = New-Object System.Windows.Forms.Label
$lblDetType.Text = "Detection Type:"; $lblDetType.AutoSize = $true
$lblDetType.Location = New-Object System.Drawing.Point(14, 12); $lblDetType.ForeColor = $clrText
$lblDetType.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
$pnlDetType.Controls.Add($lblDetType)

$cmbDetType = New-Object System.Windows.Forms.ComboBox
$cmbDetType.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
$cmbDetType.Location = New-Object System.Drawing.Point(130, 8); $cmbDetType.Size = New-Object System.Drawing.Size(200, 24)
$cmbDetType.Font = New-Object System.Drawing.Font("Segoe UI", 9)
$cmbDetType.BackColor = $clrPanelBg; $cmbDetType.ForeColor = $clrText
@('RegistryKeyValue', 'RegistryKey', 'File', 'Script') | ForEach-Object { [void]$cmbDetType.Items.Add($_) }
$cmbDetType.SelectedIndex = 0
$pnlDetType.Controls.Add($cmbDetType)

# -- Separator --
$sepInput = New-Object System.Windows.Forms.Panel
$sepInput.Dock = [System.Windows.Forms.DockStyle]::Top; $sepInput.Height = 1; $sepInput.BackColor = $clrSepLine
$tabDetection.Controls.Add($sepInput)

# -- Input panels container (fixed height, overlapping panels toggled by Visible) --
$pnlInputContainer = New-Object System.Windows.Forms.Panel
$pnlInputContainer.Dock = [System.Windows.Forms.DockStyle]::Top
$pnlInputContainer.Height = 170; $pnlInputContainer.BackColor = $clrFormBg

# Helper to create themed labels/textboxes
function New-InputLabel {
    param([string]$Text, [int]$X, [int]$Y)
    $lbl = New-Object System.Windows.Forms.Label
    $lbl.Text = $Text; $lbl.AutoSize = $true
    $lbl.Location = New-Object System.Drawing.Point($X, $Y)
    $lbl.ForeColor = $clrText; $lbl.Font = New-Object System.Drawing.Font("Segoe UI", 9)
    return $lbl
}

function New-InputTextBox {
    param([int]$X, [int]$Y, [int]$W, [int]$H = 24)
    $txt = New-Object System.Windows.Forms.TextBox
    $txt.Location = New-Object System.Drawing.Point($X, $Y); $txt.Size = New-Object System.Drawing.Size($W, $H)
    $txt.Font = New-Object System.Drawing.Font("Consolas", 9)
    $txt.BackColor = $clrPanelBg; $txt.ForeColor = $clrText
    $txt.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle
    return $txt
}

# --- Panel: RegistryKeyValue ---
$pnlRegKeyVal = New-Object System.Windows.Forms.Panel
$pnlRegKeyVal.SetBounds(0, 0, 1100, 170); $pnlRegKeyVal.BackColor = $clrFormBg
$pnlRegKeyVal.Visible = $true

$pnlRegKeyVal.Controls.Add((New-InputLabel "Registry Key:" 14 10))
$txtRegKeyVal_Key = New-InputTextBox 140 8 700
$pnlRegKeyVal.Controls.Add($txtRegKeyVal_Key)

$pnlRegKeyVal.Controls.Add((New-InputLabel "Value Name:" 14 42))
$txtRegKeyVal_ValName = New-InputTextBox 140 40 400
$txtRegKeyVal_ValName.Text = "DisplayVersion"
$pnlRegKeyVal.Controls.Add($txtRegKeyVal_ValName)

$pnlRegKeyVal.Controls.Add((New-InputLabel "Expected Value:" 14 74))
$txtRegKeyVal_Expected = New-InputTextBox 140 72 400
$pnlRegKeyVal.Controls.Add($txtRegKeyVal_Expected)

$pnlRegKeyVal.Controls.Add((New-InputLabel "Operator:" 14 106))
$cmbRegKeyVal_Op = New-Object System.Windows.Forms.ComboBox
$cmbRegKeyVal_Op.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
$cmbRegKeyVal_Op.Location = New-Object System.Drawing.Point(140, 104); $cmbRegKeyVal_Op.Size = New-Object System.Drawing.Size(150, 24)
$cmbRegKeyVal_Op.Font = New-Object System.Drawing.Font("Segoe UI", 9)
$cmbRegKeyVal_Op.BackColor = $clrPanelBg; $cmbRegKeyVal_Op.ForeColor = $clrText
@('IsEquals', 'GreaterEquals', 'GreaterThan', 'LessEquals', 'LessThan') | ForEach-Object { [void]$cmbRegKeyVal_Op.Items.Add($_) }
$cmbRegKeyVal_Op.SelectedIndex = 0
$pnlRegKeyVal.Controls.Add($cmbRegKeyVal_Op)

$pnlRegKeyVal.Controls.Add((New-InputLabel "Property:" 310 106))
$cmbRegKeyVal_PropType = New-Object System.Windows.Forms.ComboBox
$cmbRegKeyVal_PropType.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
$cmbRegKeyVal_PropType.Location = New-Object System.Drawing.Point(380, 104); $cmbRegKeyVal_PropType.Size = New-Object System.Drawing.Size(120, 24)
$cmbRegKeyVal_PropType.Font = New-Object System.Drawing.Font("Segoe UI", 9)
$cmbRegKeyVal_PropType.BackColor = $clrPanelBg; $cmbRegKeyVal_PropType.ForeColor = $clrText
@('String', 'Version', 'Integer') | ForEach-Object { [void]$cmbRegKeyVal_PropType.Items.Add($_) }
$cmbRegKeyVal_PropType.SelectedIndex = 0
$pnlRegKeyVal.Controls.Add($cmbRegKeyVal_PropType)

$chkRegKeyVal_64 = New-Object System.Windows.Forms.CheckBox
$chkRegKeyVal_64.Text = "64-bit"; $chkRegKeyVal_64.AutoSize = $true
$chkRegKeyVal_64.Location = New-Object System.Drawing.Point(520, 106); $chkRegKeyVal_64.Checked = $true
$chkRegKeyVal_64.ForeColor = $clrText; $chkRegKeyVal_64.Font = New-Object System.Drawing.Font("Segoe UI", 9)
if ($script:Prefs.DarkMode) { $chkRegKeyVal_64.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat }
$pnlRegKeyVal.Controls.Add($chkRegKeyVal_64)

$pnlInputContainer.Controls.Add($pnlRegKeyVal)

# --- Panel: RegistryKey ---
$pnlRegKey = New-Object System.Windows.Forms.Panel
$pnlRegKey.SetBounds(0, 0, 1100, 170); $pnlRegKey.BackColor = $clrFormBg
$pnlRegKey.Visible = $false

$pnlRegKey.Controls.Add((New-InputLabel "Registry Key:" 14 10))
$txtRegKey_Key = New-InputTextBox 140 8 700
$pnlRegKey.Controls.Add($txtRegKey_Key)

$chkRegKey_64 = New-Object System.Windows.Forms.CheckBox
$chkRegKey_64.Text = "64-bit"; $chkRegKey_64.AutoSize = $true
$chkRegKey_64.Location = New-Object System.Drawing.Point(140, 42); $chkRegKey_64.Checked = $true
$chkRegKey_64.ForeColor = $clrText; $chkRegKey_64.Font = New-Object System.Drawing.Font("Segoe UI", 9)
if ($script:Prefs.DarkMode) { $chkRegKey_64.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat }
$pnlRegKey.Controls.Add($chkRegKey_64)

$pnlInputContainer.Controls.Add($pnlRegKey)

# --- Panel: File ---
$pnlFile = New-Object System.Windows.Forms.Panel
$pnlFile.SetBounds(0, 0, 1100, 170); $pnlFile.BackColor = $clrFormBg
$pnlFile.Visible = $false

$pnlFile.Controls.Add((New-InputLabel "File Path:" 14 10))
$txtFile_Path = New-InputTextBox 140 8 700
$pnlFile.Controls.Add($txtFile_Path)

$pnlFile.Controls.Add((New-InputLabel "File Name:" 14 42))
$txtFile_Name = New-InputTextBox 140 40 400
$pnlFile.Controls.Add($txtFile_Name)

$pnlFile.Controls.Add((New-InputLabel "Check:" 14 74))
$cmbFile_PropType = New-Object System.Windows.Forms.ComboBox
$cmbFile_PropType.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
$cmbFile_PropType.Location = New-Object System.Drawing.Point(140, 72); $cmbFile_PropType.Size = New-Object System.Drawing.Size(150, 24)
$cmbFile_PropType.Font = New-Object System.Drawing.Font("Segoe UI", 9)
$cmbFile_PropType.BackColor = $clrPanelBg; $cmbFile_PropType.ForeColor = $clrText
@('Existence', 'Version') | ForEach-Object { [void]$cmbFile_PropType.Items.Add($_) }
$cmbFile_PropType.SelectedIndex = 0
$pnlFile.Controls.Add($cmbFile_PropType)

$lblFile_Expected = New-InputLabel "Expected Version:" 14 106
$pnlFile.Controls.Add($lblFile_Expected)
$txtFile_Expected = New-InputTextBox 140 104 300
$pnlFile.Controls.Add($txtFile_Expected)

$lblFile_Op = New-InputLabel "Operator:" 460 106
$pnlFile.Controls.Add($lblFile_Op)
$cmbFile_Op = New-Object System.Windows.Forms.ComboBox
$cmbFile_Op.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
$cmbFile_Op.Location = New-Object System.Drawing.Point(530, 104); $cmbFile_Op.Size = New-Object System.Drawing.Size(150, 24)
$cmbFile_Op.Font = New-Object System.Drawing.Font("Segoe UI", 9)
$cmbFile_Op.BackColor = $clrPanelBg; $cmbFile_Op.ForeColor = $clrText
@('GreaterEquals', 'IsEquals', 'GreaterThan', 'LessEquals', 'LessThan') | ForEach-Object { [void]$cmbFile_Op.Items.Add($_) }
$cmbFile_Op.SelectedIndex = 0
$pnlFile.Controls.Add($cmbFile_Op)

# Toggle version fields based on PropertyType
$cmbFile_PropType.Add_SelectedIndexChanged({
    $isVersion = ($cmbFile_PropType.SelectedItem -eq 'Version')
    $lblFile_Expected.Visible = $isVersion
    $txtFile_Expected.Visible = $isVersion
    $lblFile_Op.Visible = $isVersion
    $cmbFile_Op.Visible = $isVersion
}.GetNewClosure())

# Initialize visibility
$lblFile_Expected.Visible = $false; $txtFile_Expected.Visible = $false
$lblFile_Op.Visible = $false; $cmbFile_Op.Visible = $false

$pnlInputContainer.Controls.Add($pnlFile)

# --- Panel: Script ---
$pnlScript = New-Object System.Windows.Forms.Panel
$pnlScript.SetBounds(0, 0, 1100, 170); $pnlScript.BackColor = $clrFormBg
$pnlScript.Visible = $false

$pnlScript.Controls.Add((New-InputLabel "PowerShell Script:" 14 10))
$txtScript_Text = New-Object System.Windows.Forms.TextBox
$txtScript_Text.Location = New-Object System.Drawing.Point(14, 32); $txtScript_Text.Size = New-Object System.Drawing.Size(826, 126)
$txtScript_Text.Font = New-Object System.Drawing.Font("Consolas", 9)
$txtScript_Text.BackColor = $clrPanelBg; $txtScript_Text.ForeColor = $clrText
$txtScript_Text.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle
$txtScript_Text.Multiline = $true; $txtScript_Text.ScrollBars = [System.Windows.Forms.ScrollBars]::Vertical
$txtScript_Text.AcceptsReturn = $true; $txtScript_Text.AcceptsTab = $true
$pnlScript.Controls.Add($txtScript_Text)

$pnlInputContainer.Controls.Add($pnlScript)

$tabDetection.Controls.Add($pnlInputContainer)

# -- Panel swap on detection type change --
$script:InputPanels = @($pnlRegKeyVal, $pnlRegKey, $pnlFile, $pnlScript)

$cmbDetType.Add_SelectedIndexChanged({
    $idx = $cmbDetType.SelectedIndex
    for ($i = 0; $i -lt $script:InputPanels.Count; $i++) {
        $script:InputPanels[$i].Visible = ($i -eq $idx)
    }
}.GetNewClosure())

# -- Separator --
$sepButtons = New-Object System.Windows.Forms.Panel
$sepButtons.Dock = [System.Windows.Forms.DockStyle]::Top; $sepButtons.Height = 1; $sepButtons.BackColor = $clrSepLine
$tabDetection.Controls.Add($sepButtons)

# -- Buttons panel --
$pnlButtons = New-Object System.Windows.Forms.Panel
$pnlButtons.Dock = [System.Windows.Forms.DockStyle]::Top
$pnlButtons.Height = 44; $pnlButtons.BackColor = $clrFormBg
$pnlButtons.Padding = New-Object System.Windows.Forms.Padding(12, 6, 12, 6)

$btnTest = New-Object System.Windows.Forms.Button
$btnTest.Text = "Test Detection"; $btnTest.SetBounds(14, 6, 130, 30)
$btnTest.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
Set-ModernButtonStyle -Button $btnTest -BackColor $clrAccent
$pnlButtons.Controls.Add($btnTest)

$btnClear = New-Object System.Windows.Forms.Button
$btnClear.Text = "Clear"; $btnClear.SetBounds(152, 6, 80, 30)
$btnClear.Font = New-Object System.Drawing.Font("Segoe UI", 9)
$btnClear.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
$btnClear.ForeColor = $clrText; $btnClear.BackColor = $clrFormBg
$pnlButtons.Controls.Add($btnClear)

$btnImport = New-Object System.Windows.Forms.Button
$btnImport.Text = "Import Manifest..."; $btnImport.SetBounds(240, 6, 140, 30)
$btnImport.Font = New-Object System.Drawing.Font("Segoe UI", 9)
$btnImport.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
$btnImport.ForeColor = $clrText; $btnImport.BackColor = $clrFormBg
$pnlButtons.Controls.Add($btnImport)

$tabDetection.Controls.Add($pnlButtons)

# -- Separator --
$sepResults = New-Object System.Windows.Forms.Panel
$sepResults.Dock = [System.Windows.Forms.DockStyle]::Top; $sepResults.Height = 1; $sepResults.BackColor = $clrSepLine
$tabDetection.Controls.Add($sepResults)

# -- Results panel (RichTextBox, fixed height) --
$pnlResults = New-Object System.Windows.Forms.Panel
$pnlResults.Dock = [System.Windows.Forms.DockStyle]::Top
$pnlResults.Height = 100; $pnlResults.BackColor = $clrFormBg
$pnlResults.Padding = New-Object System.Windows.Forms.Padding(12, 4, 12, 4)

$rtbResults = New-Object System.Windows.Forms.RichTextBox
$rtbResults.Dock = [System.Windows.Forms.DockStyle]::Fill
$rtbResults.ReadOnly = $true; $rtbResults.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle
$rtbResults.Font = New-Object System.Drawing.Font("Consolas", 9.5)
$rtbResults.BackColor = $clrDetailBg; $rtbResults.ForeColor = $clrText
$rtbResults.Text = "Run a detection test to see results here."
$pnlResults.Controls.Add($rtbResults)

$tabDetection.Controls.Add($pnlResults)

# -- Separator --
$sepHistory = New-Object System.Windows.Forms.Panel
$sepHistory.Dock = [System.Windows.Forms.DockStyle]::Top; $sepHistory.Height = 1; $sepHistory.BackColor = $clrSepLine
$tabDetection.Controls.Add($sepHistory)

# -- History label --
$pnlHistoryLabel = New-Object System.Windows.Forms.Panel
$pnlHistoryLabel.Dock = [System.Windows.Forms.DockStyle]::Top
$pnlHistoryLabel.Height = 26; $pnlHistoryLabel.BackColor = $clrFormBg

$lblHistory = New-Object System.Windows.Forms.Label
$lblHistory.Text = "Test History"; $lblHistory.AutoSize = $true
$lblHistory.Location = New-Object System.Drawing.Point(14, 4)
$lblHistory.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
$lblHistory.ForeColor = $clrText
$pnlHistoryLabel.Controls.Add($lblHistory)

$tabDetection.Controls.Add($pnlHistoryLabel)

# -- History grid (fills remaining space) --
$dgvHistory = New-ThemedGrid
$dgvHistory.Dock = [System.Windows.Forms.DockStyle]::Fill

$colTime     = New-Object System.Windows.Forms.DataGridViewTextBoxColumn; $colTime.Name = "Time";     $colTime.DataPropertyName = "Time";     $colTime.Width = 80
$colType     = New-Object System.Windows.Forms.DataGridViewTextBoxColumn; $colType.Name = "Type";     $colType.DataPropertyName = "Type";     $colType.Width = 130
$colTarget   = New-Object System.Windows.Forms.DataGridViewTextBoxColumn; $colTarget.Name = "Target";   $colTarget.DataPropertyName = "Target";   $colTarget.AutoSizeMode = [System.Windows.Forms.DataGridViewAutoSizeColumnMode]::Fill
$colExpected = New-Object System.Windows.Forms.DataGridViewTextBoxColumn; $colExpected.Name = "Expected"; $colExpected.DataPropertyName = "Expected"; $colExpected.Width = 160
$colFound    = New-Object System.Windows.Forms.DataGridViewTextBoxColumn; $colFound.Name = "Found";    $colFound.DataPropertyName = "Found";    $colFound.Width = 160
$colResult   = New-Object System.Windows.Forms.DataGridViewTextBoxColumn; $colResult.Name = "Result";   $colResult.DataPropertyName = "Result";   $colResult.Width = 100

[void]$dgvHistory.Columns.AddRange([System.Windows.Forms.DataGridViewColumn[]]@($colTime, $colType, $colTarget, $colExpected, $colFound, $colResult))
$dgvHistory.DataSource = $script:HistoryTable

# Color-code Result column
$dgvHistory.Add_CellFormatting({
    param($s, $e)
    if ($e.ColumnIndex -eq 5 -and $null -ne $e.Value) {
        $val = [string]$e.Value
        if ($val -eq 'Detected') {
            $e.CellStyle.ForeColor = $clrOkText
            $e.CellStyle.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
        } elseif ($val -eq 'Not Detected') {
            $e.CellStyle.ForeColor = $clrErrText
            $e.CellStyle.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
        }
    }
}.GetNewClosure())

$tabDetection.Controls.Add($dgvHistory)
$dgvHistory.BringToFront()

# -- History context menu --
$ctxHistory = New-Object System.Windows.Forms.ContextMenuStrip
if ($script:Prefs.DarkMode -and $script:DarkRenderer) { $ctxHistory.Renderer = $script:DarkRenderer }

$ctxCopy = New-Object System.Windows.Forms.ToolStripMenuItem("Copy Row")
$ctxCopy.ForeColor = $clrText
$ctxCopy.Add_Click({
    if ($dgvHistory.SelectedRows.Count -gt 0) {
        $row = $dgvHistory.SelectedRows[0]
        $parts = @()
        foreach ($cell in $row.Cells) { $parts += [string]$cell.Value }
        [System.Windows.Forms.Clipboard]::SetText($parts -join "`t")
        $statusLabel.Text = "Copied row to clipboard"
    }
}.GetNewClosure())
[void]$ctxHistory.Items.Add($ctxCopy)

$ctxExportCsv = New-Object System.Windows.Forms.ToolStripMenuItem("Export CSV...")
$ctxExportCsv.ForeColor = $clrText
$ctxExportCsv.Add_Click({
    if ($script:HistoryTable.Rows.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show("No test history to export.", "Export", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information) | Out-Null
        return
    }
    $sfd = New-Object System.Windows.Forms.SaveFileDialog
    $sfd.Filter = "CSV files (*.csv)|*.csv"; $sfd.DefaultExt = "csv"
    $sfd.FileName = "DetectionTests-$(Get-Date -Format 'yyyyMMdd-HHmmss').csv"
    $sfd.InitialDirectory = Join-Path $PSScriptRoot "Reports"
    if ($sfd.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        Export-DetectionResultsCsv -DataTable $script:HistoryTable -OutputPath $sfd.FileName
        $statusLabel.Text = "Exported to $($sfd.FileName)"
    }
    $sfd.Dispose()
}.GetNewClosure())
[void]$ctxHistory.Items.Add($ctxExportCsv)

$ctxExportHtml = New-Object System.Windows.Forms.ToolStripMenuItem("Export HTML...")
$ctxExportHtml.ForeColor = $clrText
$ctxExportHtml.Add_Click({
    if ($script:HistoryTable.Rows.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show("No test history to export.", "Export", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information) | Out-Null
        return
    }
    $sfd = New-Object System.Windows.Forms.SaveFileDialog
    $sfd.Filter = "HTML files (*.html)|*.html"; $sfd.DefaultExt = "html"
    $sfd.FileName = "DetectionTests-$(Get-Date -Format 'yyyyMMdd-HHmmss').html"
    $sfd.InitialDirectory = Join-Path $PSScriptRoot "Reports"
    if ($sfd.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        Export-DetectionResultsHtml -DataTable $script:HistoryTable -OutputPath $sfd.FileName -ReportTitle "Detection Test Results"
        $statusLabel.Text = "Exported to $($sfd.FileName)"
    }
    $sfd.Dispose()
}.GetNewClosure())
[void]$ctxHistory.Items.Add($ctxExportHtml)

$dgvHistory.ContextMenuStrip = $ctxHistory

# -- Dock order: BringToFront processes last = innermost = Fill area --
# Top-docked panels process first (outermost), grid fills remaining
$pnlDetType.BringToFront()
$sepInput.BringToFront()
$pnlInputContainer.BringToFront()
$sepButtons.BringToFront()
$pnlButtons.BringToFront()
$sepResults.BringToFront()
$pnlResults.BringToFront()
$sepHistory.BringToFront()
$pnlHistoryLabel.BringToFront()
$dgvHistory.BringToFront()

# ---------------------------------------------------------------------------
# Results helper
# ---------------------------------------------------------------------------

$script:ShowDetectionResult = {
    param([hashtable]$TestResult)

    $rtbResults.Clear()

    $detected = $TestResult.Detected
    $resultText = if ($detected) { "DETECTED" } else { "NOT DETECTED" }
    $resultColor = if ($detected) { $clrOkText } else { $clrErrText }

    # Result header
    $rtbResults.SelectionFont = New-Object System.Drawing.Font("Consolas", 11, [System.Drawing.FontStyle]::Bold)
    $rtbResults.SelectionColor = $resultColor
    $rtbResults.AppendText("Result: $resultText`r`n")

    # Details
    $rtbResults.SelectionFont = New-Object System.Drawing.Font("Consolas", 9.5)
    $rtbResults.SelectionColor = $clrText
    $rtbResults.AppendText($TestResult.Details)

    if ($TestResult.Type -eq 'RegistryKeyValue') {
        $rtbResults.AppendText("`r`nKey Exists: $(if ($TestResult.KeyExists) { 'Yes' } else { 'No' })")
        if ($TestResult.ValueFound) {
            $rtbResults.AppendText(" | Value: `"$($TestResult.ActualValue)`" | Expected: `"$($TestResult.ExpectedValue)`"")
        }
    } elseif ($TestResult.Type -eq 'File') {
        $rtbResults.AppendText("`r`nFile Exists: $(if ($TestResult.FileExists) { 'Yes' } else { 'No' })")
        if ($TestResult.ActualValue) {
            $rtbResults.AppendText(" | Version: $($TestResult.ActualValue)")
        }
    } elseif ($TestResult.Type -eq 'Script' -and $TestResult.ScriptOutput) {
        $rtbResults.AppendText("`r`nOutput: $($TestResult.ScriptOutput)")
    }

    # Add to history
    $time = (Get-Date).ToString("HH:mm:ss")
    $found = if ($TestResult.ActualValue) { $TestResult.ActualValue } else { '--' }
    $expected = if ($TestResult.ExpectedValue) { $TestResult.ExpectedValue } else { '--' }

    [void]$script:HistoryTable.Rows.Add($time, $TestResult.Type, $TestResult.Target, $expected, $found, $resultText)
    $statusLabel.Text = "Test complete: $resultText"
}.GetNewClosure()

# ---------------------------------------------------------------------------
# Test button handler
# ---------------------------------------------------------------------------

$btnTest.Add_Click({
    $type = [string]$cmbDetType.SelectedItem

    switch ($type) {
        'RegistryKeyValue' {
            if ([string]::IsNullOrWhiteSpace($txtRegKeyVal_Key.Text)) {
                [System.Windows.Forms.MessageBox]::Show("Registry Key is required.", "Validation", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning) | Out-Null
                return
            }
            $p = @{
                RegistryKeyRelative = $txtRegKeyVal_Key.Text.Trim()
                ValueName           = $txtRegKeyVal_ValName.Text.Trim()
                ExpectedValue       = $txtRegKeyVal_Expected.Text.Trim()
                Operator            = [string]$cmbRegKeyVal_Op.SelectedItem
                PropertyType        = [string]$cmbRegKeyVal_PropType.SelectedItem
                Is64Bit             = $chkRegKeyVal_64.Checked
            }
            $result = Test-RegistryKeyValueDetection @p
            & $script:ShowDetectionResult $result
        }
        'RegistryKey' {
            if ([string]::IsNullOrWhiteSpace($txtRegKey_Key.Text)) {
                [System.Windows.Forms.MessageBox]::Show("Registry Key is required.", "Validation", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning) | Out-Null
                return
            }
            $p = @{
                RegistryKeyRelative = $txtRegKey_Key.Text.Trim()
                Is64Bit             = $chkRegKey_64.Checked
            }
            $result = Test-RegistryKeyDetection @p
            & $script:ShowDetectionResult $result
        }
        'File' {
            if ([string]::IsNullOrWhiteSpace($txtFile_Path.Text) -or [string]::IsNullOrWhiteSpace($txtFile_Name.Text)) {
                [System.Windows.Forms.MessageBox]::Show("File Path and File Name are required.", "Validation", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning) | Out-Null
                return
            }
            $p = @{
                FilePath     = $txtFile_Path.Text.Trim()
                FileName     = $txtFile_Name.Text.Trim()
                PropertyType = [string]$cmbFile_PropType.SelectedItem
            }
            if ($cmbFile_PropType.SelectedItem -eq 'Version') {
                $p.ExpectedValue = $txtFile_Expected.Text.Trim()
                $p.Operator = [string]$cmbFile_Op.SelectedItem
            }
            $result = Test-FileDetection @p
            & $script:ShowDetectionResult $result
        }
        'Script' {
            if ([string]::IsNullOrWhiteSpace($txtScript_Text.Text)) {
                [System.Windows.Forms.MessageBox]::Show("Script text is required.", "Validation", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning) | Out-Null
                return
            }
            $result = Test-ScriptDetection -ScriptText $txtScript_Text.Text
            & $script:ShowDetectionResult $result
        }
    }
}.GetNewClosure())

# ---------------------------------------------------------------------------
# Clear button handler
# ---------------------------------------------------------------------------

$btnClear.Add_Click({
    $type = [string]$cmbDetType.SelectedItem
    switch ($type) {
        'RegistryKeyValue' {
            $txtRegKeyVal_Key.Clear(); $txtRegKeyVal_ValName.Text = "DisplayVersion"
            $txtRegKeyVal_Expected.Clear(); $cmbRegKeyVal_Op.SelectedIndex = 0
            $cmbRegKeyVal_PropType.SelectedIndex = 0; $chkRegKeyVal_64.Checked = $true
        }
        'RegistryKey' { $txtRegKey_Key.Clear(); $chkRegKey_64.Checked = $true }
        'File' {
            $txtFile_Path.Clear(); $txtFile_Name.Clear()
            $cmbFile_PropType.SelectedIndex = 0; $txtFile_Expected.Clear(); $cmbFile_Op.SelectedIndex = 0
        }
        'Script' { $txtScript_Text.Clear() }
    }
    $rtbResults.Clear(); $rtbResults.Text = "Run a detection test to see results here."
    $statusLabel.Text = "Cleared"
}.GetNewClosure())

# ---------------------------------------------------------------------------
# Import manifest button handler
# ---------------------------------------------------------------------------

$btnImport.Add_Click({
    $ofd = New-Object System.Windows.Forms.OpenFileDialog
    $ofd.Filter = "JSON files (*.json)|*.json|All files (*.*)|*.*"
    $ofd.Title = "Import Detection Manifest"
    if ($ofd.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        try {
            $det = Import-DetectionManifest -Path $ofd.FileName

            switch ($det.Type) {
                'RegistryKeyValue' {
                    $cmbDetType.SelectedIndex = 0
                    $txtRegKeyVal_Key.Text = $det.RegistryKeyRelative
                    $txtRegKeyVal_ValName.Text = $det.ValueName
                    $txtRegKeyVal_Expected.Text = $det.ExpectedValue
                    # Find operator in combo
                    $opIdx = $cmbRegKeyVal_Op.Items.IndexOf($det.Operator)
                    if ($opIdx -ge 0) { $cmbRegKeyVal_Op.SelectedIndex = $opIdx }
                    $ptIdx = $cmbRegKeyVal_PropType.Items.IndexOf($det.PropertyType)
                    if ($ptIdx -ge 0) { $cmbRegKeyVal_PropType.SelectedIndex = $ptIdx }
                    $chkRegKeyVal_64.Checked = $det.Is64Bit
                }
                'RegistryKey' {
                    $cmbDetType.SelectedIndex = 1
                    $txtRegKey_Key.Text = $det.RegistryKeyRelative
                    $chkRegKey_64.Checked = $det.Is64Bit
                }
                'File' {
                    $cmbDetType.SelectedIndex = 2
                    $txtFile_Path.Text = $det.FilePath
                    $txtFile_Name.Text = $det.FileName
                    $ptIdx = $cmbFile_PropType.Items.IndexOf($det.PropertyType)
                    if ($ptIdx -ge 0) { $cmbFile_PropType.SelectedIndex = $ptIdx }
                    $txtFile_Expected.Text = $det.ExpectedValue
                    $opIdx = $cmbFile_Op.Items.IndexOf($det.Operator)
                    if ($opIdx -ge 0) { $cmbFile_Op.SelectedIndex = $opIdx }
                }
                'Script' {
                    $cmbDetType.SelectedIndex = 3
                    $txtScript_Text.Text = $det.ScriptText
                }
                'Compound' {
                    [System.Windows.Forms.MessageBox]::Show(
                        "Compound detection imported with $($det.Clauses.Count) clauses ($($det.Connector)).`r`nRunning all clauses now...",
                        "Compound Detection", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information
                    ) | Out-Null
                    $compResult = Test-CompoundDetection -Connector $det.Connector -Clauses $det.Clauses
                    & $script:ShowDetectionResult $compResult
                    return
                }
            }

            $statusLabel.Text = "Imported $($det.Type) detection from $(Split-Path $ofd.FileName -Leaf)"
        } catch {
            [System.Windows.Forms.MessageBox]::Show("Import failed: $_", "Import Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error) | Out-Null
        }
    }
    $ofd.Dispose()
}.GetNewClosure())

# ===========================================================================
# TAB 2: Installed Applications
# ===========================================================================

$tabApps = New-Object System.Windows.Forms.TabPage
$tabApps.Text = "Installed Applications"
$tabApps.BackColor = $clrFormBg
$tabMain.TabPages.Add($tabApps)

# -- Apps DataTable --
$script:AppsTable = New-Object System.Data.DataTable
[void]$script:AppsTable.Columns.Add("DisplayName", [string])
[void]$script:AppsTable.Columns.Add("Publisher", [string])
[void]$script:AppsTable.Columns.Add("DisplayVersion", [string])
[void]$script:AppsTable.Columns.Add("Architecture", [string])
[void]$script:AppsTable.Columns.Add("RegistryKey", [string])
[void]$script:AppsTable.Columns.Add("UninstallString", [string])
[void]$script:AppsTable.Columns.Add("QuietUninstallString", [string])
[void]$script:AppsTable.Columns.Add("InstallLocation", [string])
[void]$script:AppsTable.Columns.Add("InstallDate", [string])

$script:AppsView = New-Object System.Data.DataView($script:AppsTable)
$script:AppsLoaded = $false

# -- Filter panel --
$pnlAppsFilter = New-Object System.Windows.Forms.Panel
$pnlAppsFilter.Dock = [System.Windows.Forms.DockStyle]::Top
$pnlAppsFilter.Height = 44; $pnlAppsFilter.BackColor = $clrFormBg

$lblAppsFilter = New-Object System.Windows.Forms.Label
$lblAppsFilter.Text = "Filter:"; $lblAppsFilter.AutoSize = $true
$lblAppsFilter.Location = New-Object System.Drawing.Point(14, 12)
$lblAppsFilter.ForeColor = $clrText; $lblAppsFilter.Font = New-Object System.Drawing.Font("Segoe UI", 9)
$pnlAppsFilter.Controls.Add($lblAppsFilter)

$txtAppsFilter = New-Object System.Windows.Forms.TextBox
$txtAppsFilter.Location = New-Object System.Drawing.Point(60, 9); $txtAppsFilter.Size = New-Object System.Drawing.Size(400, 24)
$txtAppsFilter.Font = New-Object System.Drawing.Font("Segoe UI", 9)
$txtAppsFilter.BackColor = $clrPanelBg; $txtAppsFilter.ForeColor = $clrText
$txtAppsFilter.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle
$pnlAppsFilter.Controls.Add($txtAppsFilter)

$btnUseForDetection = New-Object System.Windows.Forms.Button
$btnUseForDetection.Text = "Use for Detection"; $btnUseForDetection.SetBounds(480, 7, 140, 28)
$btnUseForDetection.Font = New-Object System.Drawing.Font("Segoe UI", 9)
Set-ModernButtonStyle -Button $btnUseForDetection -BackColor $clrAccent
$pnlAppsFilter.Controls.Add($btnUseForDetection)

$lblAppsCount = New-Object System.Windows.Forms.Label
$lblAppsCount.Text = ""; $lblAppsCount.AutoSize = $true
$lblAppsCount.Location = New-Object System.Drawing.Point(636, 12)
$lblAppsCount.ForeColor = $clrHint; $lblAppsCount.Font = New-Object System.Drawing.Font("Segoe UI", 9)
$pnlAppsFilter.Controls.Add($lblAppsCount)

$tabApps.Controls.Add($pnlAppsFilter)

# -- Separator --
$sepApps = New-Object System.Windows.Forms.Panel
$sepApps.Dock = [System.Windows.Forms.DockStyle]::Top; $sepApps.Height = 1; $sepApps.BackColor = $clrSepLine
$tabApps.Controls.Add($sepApps)

# -- SplitContainer: grid (top) + detail panel (bottom) --
$splitApps = New-Object System.Windows.Forms.SplitContainer
$splitApps.Dock = [System.Windows.Forms.DockStyle]::Fill
$splitApps.Orientation = [System.Windows.Forms.Orientation]::Horizontal
$splitApps.SplitterDistance = 350
$splitApps.SplitterWidth = 6
$splitApps.BackColor = $clrSepLine
$splitApps.Panel1.BackColor = $clrPanelBg
$splitApps.Panel2.BackColor = $clrPanelBg
$splitApps.Panel1MinSize = 100
$splitApps.Panel2MinSize = 80

# -- Apps grid --
$dgvApps = New-ThemedGrid
$dgvApps.Dock = [System.Windows.Forms.DockStyle]::Fill

$colAppName    = New-Object System.Windows.Forms.DataGridViewTextBoxColumn; $colAppName.Name = "DisplayName";    $colAppName.DataPropertyName = "DisplayName";    $colAppName.AutoSizeMode = [System.Windows.Forms.DataGridViewAutoSizeColumnMode]::Fill
$colAppPub     = New-Object System.Windows.Forms.DataGridViewTextBoxColumn; $colAppPub.Name = "Publisher";      $colAppPub.DataPropertyName = "Publisher";       $colAppPub.Width = 180
$colAppVer     = New-Object System.Windows.Forms.DataGridViewTextBoxColumn; $colAppVer.Name = "Version";        $colAppVer.DataPropertyName = "DisplayVersion";  $colAppVer.Width = 130
$colAppArch    = New-Object System.Windows.Forms.DataGridViewTextBoxColumn; $colAppArch.Name = "Arch";          $colAppArch.DataPropertyName = "Architecture";   $colAppArch.Width = 60
$colAppRegKey  = New-Object System.Windows.Forms.DataGridViewTextBoxColumn; $colAppRegKey.Name = "Registry Key"; $colAppRegKey.DataPropertyName = "RegistryKey";  $colAppRegKey.Width = 280

[void]$dgvApps.Columns.AddRange([System.Windows.Forms.DataGridViewColumn[]]@($colAppName, $colAppPub, $colAppVer, $colAppArch, $colAppRegKey))
$dgvApps.DataSource = $script:AppsView

$splitApps.Panel1.Controls.Add($dgvApps)

# -- Detail panel (bottom) --
$pnlAppDetail = New-Object System.Windows.Forms.Panel
$pnlAppDetail.Dock = [System.Windows.Forms.DockStyle]::Fill
$pnlAppDetail.BackColor = $clrDetailBg
$pnlAppDetail.Padding = New-Object System.Windows.Forms.Padding(8, 4, 8, 4)

$btnCopyDetails = New-Object System.Windows.Forms.Button
$btnCopyDetails.Text = "Copy Details"
$btnCopyDetails.Dock = [System.Windows.Forms.DockStyle]::Top
$btnCopyDetails.Height = 28
$btnCopyDetails.Font = New-Object System.Drawing.Font("Segoe UI", 8.5, [System.Drawing.FontStyle]::Bold)
Set-ModernButtonStyle -Button $btnCopyDetails -BackColor ([System.Drawing.Color]::FromArgb(34, 139, 34))
$pnlAppDetail.Controls.Add($btnCopyDetails)

$rtbAppDetail = New-Object System.Windows.Forms.RichTextBox
$rtbAppDetail.Dock = [System.Windows.Forms.DockStyle]::Fill
$rtbAppDetail.ReadOnly = $true
$rtbAppDetail.BorderStyle = [System.Windows.Forms.BorderStyle]::None
$rtbAppDetail.Font = New-Object System.Drawing.Font("Consolas", 9.5)
$rtbAppDetail.BackColor = $clrDetailBg
$rtbAppDetail.ForeColor = $clrText
$rtbAppDetail.Text = "Select an application above to see details."
$pnlAppDetail.Controls.Add($rtbAppDetail)
$rtbAppDetail.BringToFront()

$splitApps.Panel2.Controls.Add($pnlAppDetail)

$tabApps.Controls.Add($splitApps)

# Dock order
$pnlAppsFilter.BringToFront()
$sepApps.BringToFront()
$splitApps.BringToFront()

# -- Filter debounce timer --
$script:FilterTimer = New-Object System.Windows.Forms.Timer
$script:FilterTimer.Interval = 300

$script:FilterTimer.Add_Tick({
    $script:FilterTimer.Stop()
    $filterText = $txtAppsFilter.Text.Trim()
    if ([string]::IsNullOrWhiteSpace($filterText)) {
        $script:AppsView.RowFilter = ""
    } else {
        $escaped = $filterText.Replace("'", "''").Replace("[", "[[]").Replace("%", "[%]").Replace("*", "[*]")
        $script:AppsView.RowFilter = "DisplayName LIKE '%$escaped%'"
    }
    $lblAppsCount.Text = "$($script:AppsView.Count) of $($script:AppsTable.Rows.Count) apps"
}.GetNewClosure())

$txtAppsFilter.Add_TextChanged({
    $script:FilterTimer.Stop()
    $script:FilterTimer.Start()
}.GetNewClosure())

# -- Detail panel: populate on selection change --
$script:FormatAppDetail = {
    if ($dgvApps.SelectedRows.Count -eq 0) {
        $rtbAppDetail.Text = "Select an application above to see details."
        return
    }

    $row = $dgvApps.SelectedRows[0]
    $dn = [string]$row.Cells["DisplayName"].Value
    $pub = [string]$row.Cells["Publisher"].Value
    $ver = [string]$row.Cells["Version"].Value
    $arch = [string]$row.Cells["Arch"].Value
    $regKey = [string]$row.Cells["Registry Key"].Value

    # Read hidden DataTable columns via the underlying DataRowView
    $rowIdx = $row.Index
    $drv = $script:AppsView[$rowIdx]
    $uninstall = [string]$drv["UninstallString"]
    $quietUninstall = [string]$drv["QuietUninstallString"]
    $installLoc = [string]$drv["InstallLocation"]
    $installDate = [string]$drv["InstallDate"]

    $rtbAppDetail.Clear()

    $rtbAppDetail.SelectionFont = New-Object System.Drawing.Font("Consolas", 10, [System.Drawing.FontStyle]::Bold)
    $rtbAppDetail.SelectionColor = $clrAccent
    $rtbAppDetail.AppendText("$dn`r`n")

    $rtbAppDetail.SelectionFont = New-Object System.Drawing.Font("Consolas", 9.5)
    $rtbAppDetail.SelectionColor = $clrText

    $lines = @(
        "DisplayName:          $dn"
        "Publisher:            $pub"
        "DisplayVersion:       $ver"
        "Architecture:         $arch"
        "RegistryKey:          HKLM\$regKey"
        "UninstallString:      $uninstall"
        "QuietUninstallString: $quietUninstall"
        "InstallLocation:      $installLoc"
        "InstallDate:          $installDate"
    )
    $rtbAppDetail.AppendText(($lines -join "`r`n"))
}.GetNewClosure()

$dgvApps.Add_SelectionChanged({ & $script:FormatAppDetail })

# -- Copy Details button --
$btnCopyDetails.Add_Click({
    if ($dgvApps.SelectedRows.Count -eq 0) { return }

    $row = $dgvApps.SelectedRows[0]
    $rowIdx = $row.Index
    $drv = $script:AppsView[$rowIdx]

    $lines = @(
        "DisplayName:          $([string]$drv['DisplayName'])"
        "Publisher:            $([string]$drv['Publisher'])"
        "DisplayVersion:       $([string]$drv['DisplayVersion'])"
        "Architecture:         $([string]$drv['Architecture'])"
        "RegistryKey:          HKLM\$([string]$drv['RegistryKey'])"
        "UninstallString:      $([string]$drv['UninstallString'])"
        "QuietUninstallString: $([string]$drv['QuietUninstallString'])"
        "InstallLocation:      $([string]$drv['InstallLocation'])"
        "InstallDate:          $([string]$drv['InstallDate'])"
    )
    [System.Windows.Forms.Clipboard]::SetText($lines -join "`r`n")
    $statusLabel.Text = "Copied details for: $([string]$drv['DisplayName'])"
}.GetNewClosure())

# -- Lazy-load apps on tab activation --
$tabMain.Add_SelectedIndexChanged({
    if ($tabMain.SelectedIndex -eq 1 -and -not $script:AppsLoaded) {
        $statusLabel.Text = "Loading installed applications..."
        $form.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
        try {
            $apps = Get-InstalledApplications
            $script:AppsTable.BeginLoadData()
            foreach ($app in $apps) {
                [void]$script:AppsTable.Rows.Add(
                    $app.DisplayName,
                    $app.Publisher,
                    $app.DisplayVersion,
                    $app.Architecture,
                    $app.RegistryKey,
                    $app.UninstallString,
                    $app.QuietUninstallString,
                    $app.InstallLocation,
                    $app.InstallDate
                )
            }
            $script:AppsTable.EndLoadData()
            $script:AppsLoaded = $true
            $lblAppsCount.Text = "$($script:AppsTable.Rows.Count) apps"
            $statusLabel.Text = "Loaded $($script:AppsTable.Rows.Count) installed applications"
        } catch {
            $statusLabel.Text = "Error loading applications: $_"
            Write-Log "Apps load error: $_" -Level ERROR
        }
        $form.Cursor = [System.Windows.Forms.Cursors]::Default
    }
}.GetNewClosure())

# -- Use for Detection button --
$script:UseForDetectionAction = {
    if ($dgvApps.SelectedRows.Count -eq 0) { return }

    $row = $dgvApps.SelectedRows[0]
    $regKey = [string]$row.Cells["Registry Key"].Value
    $version = [string]$row.Cells["Version"].Value

    # Switch to Detection Tester tab, set to RegistryKeyValue
    $tabMain.SelectedIndex = 0
    $cmbDetType.SelectedIndex = 0

    $txtRegKeyVal_Key.Text = $regKey
    $txtRegKeyVal_ValName.Text = "DisplayVersion"
    $txtRegKeyVal_Expected.Text = $version
    $cmbRegKeyVal_Op.SelectedIndex = 0
    $chkRegKeyVal_64.Checked = $true

    $statusLabel.Text = "Populated detection from: $([string]$row.Cells['DisplayName'].Value)"
}.GetNewClosure()

$btnUseForDetection.Add_Click($script:UseForDetectionAction)

$dgvApps.Add_CellDoubleClick({
    param($s, $e)
    if ($e.RowIndex -ge 0) { & $script:UseForDetectionAction }
}.GetNewClosure())

# ---------------------------------------------------------------------------
# Form lifecycle
# ---------------------------------------------------------------------------

$form.Add_FormClosing({
    Save-WindowState
    $script:FilterTimer.Dispose()
})

$form.Add_Shown({ Restore-WindowState })

[System.Windows.Forms.Application]::Run($form)
