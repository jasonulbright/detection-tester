# Lab-side runner. The scheduled task starts it in the interactive session. It runs steps.ps1
# and writes the finished flag: 'finished', or 'failed: <reason>' when the steps stopped.
[Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseDeclaredVarsMoreThanAssignments', '', Justification = 'steps.ps1 is dot-sourced and reads appDir and Capture.')]
param([string]$GuideRoot)
$ErrorActionPreference = 'Stop'
$OutDir = Join-Path $GuideRoot 'ui\out'; New-Item -ItemType Directory $OutDir -Force | Out-Null
$appDir = $GuideRoot
# Record this runner in the ledger, so that the capture tool stops it by process id and never by a command-line pattern.
$me = Get-Process -Id $PID
('{0}|{1}' -f $me.Id, $me.StartTime.Ticks) | Add-Content -LiteralPath (Join-Path $PSScriptRoot 'run.pids')
$flag = 'finished'
try {
    . (Join-Path $GuideRoot 'ui\guilib.ps1')
    # Settings that Invoke-GuideCapture.ps1 passes in (the sections to run).
    $Capture = Get-Content -LiteralPath (Join-Path $GuideRoot 'ui\capture.json') -Raw | ConvertFrom-Json
    . (Join-Path $GuideRoot 'ui\steps.ps1')
}
catch {
    $flag = 'failed: ' + $_.Exception.Message
    try { Log ('ERROR: ' + $_.Exception.Message + ' @ ' + $_.ScriptStackTrace) } catch { $null = $_ }
}
finally { $flag | Set-Content (Join-Path $OutDir 'finished.flag') }
