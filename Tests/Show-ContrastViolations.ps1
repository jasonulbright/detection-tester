<#
.SYNOPSIS
    Prints the contrast violations of the main window and the Options dialog in both
    themes, in the format of KnownContrastDefects.ps1.

.DESCRIPTION
    Run it when the window test reports a new or a fixed violation, to see the
    current list. Opens the windows off screen at 1440x900, as Window.Tests.ps1 does.
#>
[Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSAvoidUsingWriteHost', '', Justification='Console tool.')]
param()
. "$PSScriptRoot\TestHelpers.ps1"
. "$PSScriptRoot\WindowHarness.ps1"
. "$PSScriptRoot\WindowConfig.ps1"
. "$PSScriptRoot\KnownContrastDefects.ps1"
$tempDir = Join-Path ([System.IO.Path]::GetTempPath()) ('dt-contrast-' + [guid]::NewGuid().ToString('N'))
New-Item -ItemType Directory -Path $tempDir | Out-Null
try {
    $shellText = Get-Content -LiteralPath (Join-Path $script:AppRoot $script:AppConfig.Shell) -Raw -Encoding UTF8
    $ast = Get-ShellAst
    $dialogXaml = @($ast.FindAll({ param($n) $n -is [System.Management.Automation.Language.StringConstantExpressionAst] -and $n.Value -match '^\s*<Controls:MetroWindow' }, $true) | ForEach-Object { $_.Value })
    $colors = @{}
    foreach ($k in 'DarkButtonBg', 'DarkButtonBorder', 'LightButtonBg', 'LightButtonBorder', 'TitleBarBlue', 'TitleBarBlueInactive') {
        $colors[$k] = [regex]::Match($shellText, ('\$script:{0}\s*=\s*\[Windows\.Media\.BrushConverter\]::new\(\)\.ConvertFrom\(''(#[0-9A-Fa-f]{{6}})''\)' -f $k)).Groups[1].Value
    }
    $log = [regex]::Match($shellText, '\$hex\s*=\s*if\s*\(\$IsDark\)\s*\{\s*''(#[0-9A-Fa-f]{6})''\s*\}\s*else\s*\{\s*''(#[0-9A-Fa-f]{6})''\s*\}')
    $colors.LogDark = $log.Groups[1].Value; $colors.LogLight = $log.Groups[2].Value
    foreach ($dark in $true, $false) {
        $r = Invoke-WpfSta -Arguments @{ Dark = $dark; W = 1440; H = 900; Config = $script:AppConfig; Colors = $colors; DialogXaml = $dialogXaml[0]; Parts = (Get-ShellRuntimePart); TempDir = $tempDir } -Script $script:MeasureAppScript
        $keys = Get-ContrastViolationKey -Views @($r)
        Write-Host ("== {0}" -f $(if ($dark) { 'dark' } else { 'light' }))
        Write-Host (($keys | ForEach-Object { "'" + ($_ -replace "'", "''") + "'" }) -join ', ')
    }
}
finally { Close-WpfSession; Remove-Item -LiteralPath $tempDir -Recurse -Force -ErrorAction SilentlyContinue }
