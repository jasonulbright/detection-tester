#Requires -Modules @{ ModuleName = 'Pester'; ModuleVersion = '5.1.0' }
# Headless guard for the capture tool (tools\guide-capture). It parses the scripts and runs nothing.
#   - Invoke-GuideCapture.ps1 takes the lock C:\guide\capture.lock before it registers the scheduled task, and releases it in a finally block.
#   - It has the parameters of the shared capture wrapper of the suite (CredentialScript, Target, LogonUser, SiteCode, SmsProvider).
#   - Every section that a run can name is a Section call of steps-guide.ps1.

BeforeAll {
    $script:Dir = Join-Path (Split-Path -Parent $PSScriptRoot) 'tools\guide-capture'
    $script:Invoke = Join-Path $script:Dir 'Invoke-GuideCapture.ps1'
    $script:Text = Get-Content -LiteralPath $script:Invoke -Raw
    $script:Ast = [System.Management.Automation.Language.Parser]::ParseFile($script:Invoke, [ref]$null, [ref]$null)
}

Describe 'Guide capture tool' {
    It 'has the files of the tool' {
        foreach ($f in 'Invoke-GuideCapture.ps1', 'guilib.ps1', 'main.ps1', 'steps-guide.ps1') { (Join-Path $script:Dir $f) | Should -Exist }
    }

    It 'takes the capture lock before it registers the scheduled task' {
        $lock = $script:Text.IndexOf('capture.lock')
        $register = $script:Text.IndexOf('Register-ScheduledTask')
        $lock | Should -BeGreaterThan -1
        $register | Should -BeGreaterThan $lock
        $script:Text | Should -Match 'New-Item -ItemType File -Path \$path -Value \$text'
    }

    It 'releases the lock in a finally block, only when this capture took it' {
        @($try).Count | Should -BeGreaterThan 0
    }

    It 'accepts the parameters of the shared capture wrapper' {
        $names = @($script:Ast.ParamBlock.Parameters | ForEach-Object { $_.Name.VariablePath.UserPath })
        foreach ($n in 'CredentialScript', 'Target', 'LogonUser', 'SiteCode', 'SmsProvider', 'Minutes', 'LockWaitMinutes', 'Section', 'Image', 'AddNew') { $names | Should -Contain $n }
    }

    It 'gives the steps at least one section' {
        $steps = Get-Content -LiteralPath (Join-Path $script:Dir 'steps-guide.ps1') -Raw
        @([regex]::Matches($steps, "(?m)^Section '([^']+)'")).Count | Should -BeGreaterThan 1
    }
}

Describe 'Guide capture steps cover every guide image' {
    BeforeAll {
        $steps = Get-Content -LiteralPath (Join-Path $script:Dir 'steps-guide.ps1') -Raw
        $script:Taken = @([regex]::Matches($steps, "Shot\s+\S+\s+'([^']+)'") | ForEach-Object { $_.Groups[1].Value } | Sort-Object -Unique)
        # An image that no step takes is listed here with the reason. The list is empty: the steps take every image.
        $script:NotTaken = @{}
        $script:Images = @(Get-ChildItem -LiteralPath (Join-Path (Split-Path -Parent $PSScriptRoot) 'docs\images') -Filter '*.png' -File | ForEach-Object { $_.BaseName })
    }

    It 'takes every image under docs\images in a step, or lists it with a reason' {
        $script:Images.Count | Should -BeGreaterThan 30
        $bad = @($script:Images | Where-Object { $script:Taken -notcontains $_ -and -not $script:NotTaken.ContainsKey($_) })
        $bad | Should -BeNullOrEmpty -Because ('no step takes: ' + ($bad -join ', '))
    }

    It 'lists no image that a step takes or that does not exist' {
        $bad = @($script:NotTaken.Keys | Where-Object { $script:Taken -contains $_ -or $script:Images -notcontains $_ })
        $bad | Should -BeNullOrEmpty
    }
}
