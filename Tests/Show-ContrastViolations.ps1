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
    $ast = Get-ShellAst
    $dialogXaml = @($ast.FindAll({ param($n) $n -is [System.Management.Automation.Language.StringConstantExpressionAst] -and $n.Value -match '^\s*<Controls:MetroWindow' }, $true) | ForEach-Object { $_.Value })
    foreach ($dark in $true, $false) {
        $r = Invoke-WpfSta -Arguments @{ Dark = $dark; W = 1440; H = 900; Config = $script:AppConfig; DialogXaml = $dialogXaml[0]; Parts = (Get-ShellRuntimePart); TempDir = $tempDir } -Script $script:MeasureAppScript
        $keys = Get-ContrastViolationKey -Views @($r)
        Write-Host ("== {0}" -f $(if ($dark) { 'dark' } else { 'light' }))
        Write-Host (($keys | ForEach-Object { "'" + ($_ -replace "'", "''") + "'" }) -join ', ')
    }
}
finally { Close-WpfSession; Remove-Item -LiteralPath $tempDir -Recurse -Force -ErrorAction SilentlyContinue }
