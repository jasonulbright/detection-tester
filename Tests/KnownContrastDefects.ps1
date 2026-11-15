# Contrast violations that exist today and are not fixed yet. Each entry is
# "<control>:<metric>" or "text:<visible text>". The window test fails when a
# violation is not listed (a new defect) and when a listed one no longer happens
# (remove it here). Run Show-ContrastViolations.ps1 to print the current list.
#
# Text must reach 4.5:1 and a control border 3:1 when the control is enabled, in both themes. A disabled control is exempt:
# WCAG 1.4.3 (Contrast Minimum) exempts inactive user interface components, so the disabled text and the disabled border
# have no target. The shared button templates of SuiteCommon dim a disabled button by opacity 0.5. The window test requires
# that a disabled control stays clearly dimmed (see the test Disabled controls in Window.Tests.ps1).
# The known defect that remains comes from the SuiteCommon palette: btnOK:EnabledText and text:OK (light theme), the plain accent
# fill of the dialog primary button with white text at 4.49:1. It goes when SuiteCommon fixes it and the vendored copy is synced.
function Get-KnownContrastDefect {
    param([Parameter(Mandatory)][ValidateSet('dark', 'light')][string]$Theme)
    switch ($Theme) {
        'dark' { @() }
        'light' { @('btnOK:EnabledText', 'text:OK') }
    }
}
function Test-LargeText {
    # WCAG large text: 24 px, or 18.66 px and bold. It needs 3:1; other text needs 4.5:1.
    param([Parameter(Mandatory)]$Row)
    return ($Row.Size -ge 24 -or ($Row.Bold -and $Row.Size -ge 18.66))
}

function Get-ContrastViolationKey {
    # The one place that decides what a contrast violation is and how it is named. The window
    # test and Show-ContrastViolations.ps1 both call it, so the printed list always matches
    # what the test expects. $Views are measured views (Text, Interactive and TitleBar).
    param([Parameter(Mandatory)][AllowEmptyCollection()][object[]]$Views)
    $keys = New-Object System.Collections.Generic.List[string]
    foreach ($v in $Views) {
        foreach ($r in $v.Text) {
            $min = if (Test-LargeText $r) { 3.0 } else { 4.5 }
            if ($r.Ratio -lt $min) { [void]$keys.Add(('text:{0}' -f $r.Text)) }
        }
        foreach ($c in $v.Interactive) {
            # A disabled control is exempt from the text and border targets (WCAG 1.4.3, inactive components). Its dimming is checked apart.
            foreach ($metric in 'EnabledText', 'EnabledBorder') {
                $value = $c.$metric
                if ($null -eq $value) { continue }
                $min = if ($metric -like '*Text') { 4.5 } else { 3.0 }
                if ($value -lt $min) { [void]$keys.Add(('{0}:{1}' -f $c.Name, $metric)) }
            }
        }
    }
    $inactive = @($Views)[0].TitleBar.Inactive
    if ($null -ne $inactive -and $inactive -lt 4.5) { [void]$keys.Add('titlebar:inactive') }
    return @($keys | Sort-Object -Unique)
}
