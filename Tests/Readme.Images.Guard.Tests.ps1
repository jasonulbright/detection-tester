#Requires -Modules @{ ModuleName = 'Pester'; ModuleVersion = '5.1.0' }
# Headless guard for the images that the README and the guide show. It parses text and compares file names only.
#   - Every image that README.md shows is a file under docs\images.
#   - Every image that a guide page shows is a file under docs\images.
#   - Every image that the capture steps take (tools\guide-capture\steps-guide.ps1) is a file under docs\images,
#     so that a run of Invoke-GuideCapture.ps1 replaces it and no image is left that nothing retakes by mistake.

BeforeAll {
    $script:Root = Split-Path -Parent $PSScriptRoot
    function Get-ImageReference {
        param([string]$Path)
        $text = Get-Content -LiteralPath $Path -Raw -Encoding UTF8
        foreach ($m in [regex]::Matches($text, '!\[[^\]]*\]\(([^)\s]+\.(?:png|gif|jpg))\)')) {
            [pscustomobject]@{ Page = $Path; Ref = $m.Groups[1].Value; Full = [IO.Path]::GetFullPath((Join-Path (Split-Path -Parent $Path) $m.Groups[1].Value)) }
        }
    }
}

Describe 'README and guide images' {
    It 'shows in README.md only images from docs\images, and the files exist' {
        $refs = @(Get-ImageReference -Path (Join-Path $script:Root 'README.md'))
        $refs.Count | Should -BeGreaterThan 0
        $imageDir = [IO.Path]::GetFullPath((Join-Path $script:Root 'docs\images'))
        $bad = @($refs | Where-Object { -not $_.Full.StartsWith($imageDir + '\', [StringComparison]::OrdinalIgnoreCase) -or -not (Test-Path -LiteralPath $_.Full) } | ForEach-Object { $_.Ref })
        $bad | Should -BeNullOrEmpty
    }

    It 'shows in every guide page only images from docs\images, and the files exist' {
        $imageDir = [IO.Path]::GetFullPath((Join-Path $script:Root 'docs\images'))
        $refs = @(Get-ChildItem -LiteralPath (Join-Path $script:Root 'docs') -Filter '*.md' -Recurse -File | ForEach-Object { Get-ImageReference -Path $_.FullName })
        $refs.Count | Should -BeGreaterThan 10
        $bad = @($refs | Where-Object { -not $_.Full.StartsWith($imageDir + '\', [StringComparison]::OrdinalIgnoreCase) -or -not (Test-Path -LiteralPath $_.Full) } | ForEach-Object { '{0} -> {1}' -f (Split-Path -Leaf $_.Page), $_.Ref })
        $bad | Should -BeNullOrEmpty
    }

    It 'takes in the capture steps only images that exist in docs\images' {
        $steps = Get-Content -LiteralPath (Join-Path $script:Root 'tools\guide-capture\steps-guide.ps1') -Raw
        $names = @([regex]::Matches($steps, "Shot\s+\S+\s+'([^']+)'") | ForEach-Object { $_.Groups[1].Value } | Sort-Object -Unique)
        $names.Count | Should -BeGreaterThan 5
        $bad = @($names | Where-Object { -not (Test-Path -LiteralPath (Join-Path $script:Root ('docs\images\' + $_ + '.png'))) })
        $bad | Should -BeNullOrEmpty
    }
}
