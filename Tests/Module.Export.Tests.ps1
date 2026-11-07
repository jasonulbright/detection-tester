#Requires -Modules Pester
# CSV and HTML export.

BeforeAll {
    . "$PSScriptRoot\TestHelpers.ps1"
    Import-AppModule
    Mock Write-Log {} -ModuleName DetectionTesterCommon
}
Describe 'Codex review fixes' {
    Context "HTML report title escaping" {
        It "encodes the ReportTitle parameter" {
            $dt = New-Object System.Data.DataTable
            foreach ($c in @('Time','Type','Mode','Target','Expected','Found','Result')) {
                [void]$dt.Columns.Add($c, [string])
            }
            $row = $dt.NewRow()
            $row['Time']='00:00:00'; $row['Type']='Script'; $row['Mode']='Positive'
            $row['Target']='x'; $row['Expected']='y'; $row['Found']='z'; $row['Result']='Pass'
            [void]$dt.Rows.Add($row)
            $out = Join-Path $TestDrive "title-escape.html"
            $injected = '<img src=x onerror=alert(1)>'
            Export-DetectionResultsHtml -DataTable $dt -OutputPath $out -ReportTitle $injected
            $html = Get-Content -LiteralPath $out -Raw
            $html | Should -Not -Match '<img src=x onerror=alert'
            $html | Should -Match '&lt;img'
        }
    }

    Context "HTML export escaping" {
        It "encodes special characters in cell values" {
            $dt = New-Object System.Data.DataTable
            foreach ($c in @('Time','Type','Mode','Target','Expected','Found','Result')) {
                [void]$dt.Columns.Add($c, [string])
            }
            $row = $dt.NewRow()
            $row['Time']='00:00:00'; $row['Type']='Script'; $row['Mode']='Positive'
            $row['Target']='<img src=x onerror=alert(1)>'
            $row['Expected']='a & b'
            $row['Found']='1 < 2'
            $row['Result']='Pass'
            [void]$dt.Rows.Add($row)
            $out = Join-Path $TestDrive "html-escape.html"
            Export-DetectionResultsHtml -DataTable $dt -OutputPath $out
            $html = Get-Content -LiteralPath $out -Raw
            $html | Should -Not -Match '<img src=x onerror=alert'
            $html | Should -Match '&lt;img'
            $html | Should -Match 'a &amp; b'
            $html | Should -Match '1 &lt; 2'
        }
    }
}

Describe 'Export-DetectionResultsCsv' {
    BeforeAll {
        function script:New-Table { param([object[]]$Cell) $t = New-ResultTable; [void]$t.Rows.Add($Cell); return , $t }
    }

    It 'writes the header and the rows so that Import-Csv reads them back' {
        $t = New-ResultTable
        [void]$t.Rows.Add('10:00:01', 'File', 'Positive', 'C:\a\b.exe', '(file exists)', '', 'Pass')
        [void]$t.Rows.Add('10:00:02', 'RegistryKey', 'Negative', 'HKLM\SOFTWARE\X', '(key exists)', '', 'Fail')
        $out = Join-Path $TestDrive 'plain.csv'
        Export-DetectionResultsCsv -DataTable $t -OutputPath $out
        $rows = @(Import-Csv -LiteralPath $out -Encoding UTF8)
        $rows.Count | Should -Be 2
        ($rows[0].PSObject.Properties.Name -join ',') | Should -Be 'Time,Type,Mode,Target,Expected,Found,Result'
        $rows[0].Target | Should -Be 'C:\a\b.exe'
        $rows[1].Result | Should -Be 'Fail'
        $rows[1].Found | Should -Be ''
    }

    It 'keeps commas, quotes, line breaks and non-ASCII text in a cell' {
        $text = "a, b ""quoted""`r`nsecond line $([char]0xE9) $([char]0x20AC)"
        $out = Join-Path $TestDrive 'special.csv'
        Export-DetectionResultsCsv -DataTable (New-Table @('t', 'Script', 'Positive', '(script)', 'x', $text, 'Pass')) -OutputPath $out
        (Import-Csv -LiteralPath $out -Encoding UTF8).Found | Should -Be $text
    }

    It 'creates the folder of the output file' {
        $out = Join-Path $TestDrive 'new folder\deeper\out.csv'
        Export-DetectionResultsCsv -DataTable (New-Table @('t', 'File', 'Positive', 'a', 'b', 'c', 'Pass')) -OutputPath $out
        Test-Path -LiteralPath $out | Should -BeTrue
    }

    It 'replaces an existing file' {
        $out = Join-Path $TestDrive 'replace.csv'
        Set-Content -LiteralPath $out -Value 'old content that is much longer than the new file will be, with more words'
        Export-DetectionResultsCsv -DataTable (New-Table @('t', 'File', 'Positive', 'a', 'b', 'c', 'Pass')) -OutputPath $out
        (Get-Content -LiteralPath $out -Raw) | Should -Not -Match 'old content'
        @(Import-Csv -LiteralPath $out -Encoding UTF8).Count | Should -Be 1
    }

    It 'writes to a path with brackets' {
        $out = Join-Path $TestDrive 'a[1]\out[1].csv'
        Export-DetectionResultsCsv -DataTable (New-Table @('t', 'File', 'Positive', 'a', 'b', 'c', 'Pass')) -OutputPath $out
        Test-Path -LiteralPath $out | Should -BeTrue
    }

    It 'writes an empty cell for a database null' {
        $t = New-ResultTable
        $row = $t.NewRow(); $row['Time'] = 't'; $row['Result'] = 'Pass'
        [void]$t.Rows.Add($row)
        $out = Join-Path $TestDrive 'null.csv'
        Export-DetectionResultsCsv -DataTable $t -OutputPath $out
        $r = @(Import-Csv -LiteralPath $out -Encoding UTF8)
        $r[0].Target | Should -Be ''
        $r[0].Result | Should -Be 'Pass'
    }

    It 'does not fail for a table without rows' {
        $out = Join-Path $TestDrive 'norows.csv'
        { Export-DetectionResultsCsv -DataTable (New-ResultTable) -OutputPath $out } | Should -Not -Throw
        Test-Path -LiteralPath $out | Should -BeTrue
    }

    It 'throws when the file cannot be written: <Case>' -ForEach @(
        @{ Case = 'the folder is a file'; Kind = 'FolderIsFile' }
        @{ Case = 'the name has illegal characters'; Kind = 'IllegalName' }
        @{ Case = 'the path is a folder'; Kind = 'PathIsFolder' }
    ) {
        $out = switch ($Kind) {
            'FolderIsFile' { $f = Join-Path $TestDrive 'iamafile'; Set-Content -LiteralPath $f -Value 'x'; Join-Path $f 'out.csv' }
            'IllegalName' { Join-Path $TestDrive 'bad|name?.csv' }
            'PathIsFolder' { $f = Join-Path $TestDrive 'iamafolder.csv'; New-Item -ItemType Directory -Path $f -Force | Out-Null; $f }
        }
        { Export-DetectionResultsCsv -DataTable (New-Table @('t', 'File', 'Positive', 'a', 'b', 'c', 'Pass')) -OutputPath $out } | Should -Throw
    }

    It 'logs the export once with the path' {
        $out = Join-Path $TestDrive 'logged.csv'
        Export-DetectionResultsCsv -DataTable (New-Table @('t', 'File', 'Positive', 'a', 'b', 'c', 'Pass')) -OutputPath $out
        Should -Invoke Write-Log -ModuleName DetectionTesterCommon -ParameterFilter { $Message -eq "Exported CSV to $out" } -Times 1 -Exactly
    }

    # A cell that starts with = + - @ (or a tab or a carriage return) runs as a formula when someone opens the file
    # in a spreadsheet. The cells hold registry data and script output, which the user does not control. The export
    # puts a single quote in front of such a cell. A number such as -5 is not a formula and stays as it is.
    It 'prefixes the cell "<Cell>" with a single quote when it can run as a formula' -ForEach @(
        @{ Cell = '=1+1' }
        @{ Cell = '=cmd|'' /C calc''!A0' }
        @{ Cell = '@SUM(A1:A9)' }
        @{ Cell = '+cmd|''/c calc''!A1' }
        @{ Cell = '-2+3' }
        @{ Cell = '-' }
        @{ Cell = '=' }
        @{ Cell = "`t=1+1" }
        @{ Cell = "`r=1+1" }
    ) {
        $out = Join-Path $TestDrive 'formula.csv'
        Export-DetectionResultsCsv -DataTable (New-Table @('t', 'Script', 'Positive', $Cell, $Cell, $Cell, 'Pass')) -OutputPath $out
        $r = @(Import-Csv -LiteralPath $out -Encoding UTF8)[0]
        $r.Target | Should -Be ("'" + $Cell)
        $r.Expected | Should -Be ("'" + $Cell)
        $r.Found | Should -Be ("'" + $Cell)
    }

    It 'leaves the cell "<Cell>" as it is' -ForEach @(
        @{ Cell = '-5' }
        @{ Cell = '+5' }
        @{ Cell = '-1.5e3' }
        @{ Cell = '-0' }
        @{ Cell = 'plain text' }
        @{ Cell = 'a=b' }
        @{ Cell = '1.2.3' }
        @{ Cell = 'x@y.com' }
        @{ Cell = '' }
        @{ Cell = ' =1' }
    ) {
        $out = Join-Path $TestDrive 'safe.csv'
        Export-DetectionResultsCsv -DataTable (New-Table @('t', 'Script', 'Positive', $Cell, $Cell, $Cell, 'Pass')) -OutputPath $out
        $r = @(Import-Csv -LiteralPath $out -Encoding UTF8)[0]
        $r.Target | Should -Be $Cell
        $r.Found | Should -Be $Cell
    }

    It 'protects a column name that a spreadsheet would run as a formula' {
        $t = New-ResultTable -Column @('=Time', 'Type')
        [void]$t.Rows.Add('a', 'b')
        $out = Join-Path $TestDrive 'header.csv'
        Export-DetectionResultsCsv -DataTable $t -OutputPath $out
        (Import-Csv -LiteralPath $out -Encoding UTF8)[0].PSObject.Properties.Name | Should -Contain "'=Time"
    }
}

Describe 'Export-DetectionResultsHtml' {
    BeforeAll {
        function script:New-Table { param([object[]]$Cell) $t = New-ResultTable; [void]$t.Rows.Add($Cell); return , $t }
        function script:Read-Html { param([string]$Path) [System.IO.File]::ReadAllText($Path, [System.Text.Encoding]::UTF8) }
    }

    It 'writes a complete page with the title, the summary, the header and one row per result' {
        $t = New-ResultTable
        [void]$t.Rows.Add('10:00:01', 'File', 'Positive', 'C:\a\b.exe', '(file exists)', '', 'Pass')
        [void]$t.Rows.Add('10:00:02', 'RegistryKey', 'Negative', 'HKLM\SOFTWARE\X', '(key exists)', '', 'Fail')
        $out = Join-Path $TestDrive 'page.html'
        Export-DetectionResultsHtml -DataTable $t -OutputPath $out -ReportTitle 'My report'
        $html = Read-Html $out
        $html | Should -Match '^<!DOCTYPE html>'
        $html | Should -Match '<title>My report</title>'
        $html | Should -Match '<h1>My report</h1>'
        $html | Should -Match 'Tests: 2'
        $html | Should -Match 'Generated: \d{4}-\d{2}-\d{2} \d{2}:\d{2}:\d{2}'
        ([regex]::Matches($html, '<th>')).Count | Should -Be 7
        ([regex]::Matches($html, '<tr>')).Count | Should -Be 3
        $html | Should -Match '<td>C:\\a\\b.exe</td>'
        $html | Should -Match '</html>\s*$'
    }

    It 'uses the default title when none is given' {
        $out = Join-Path $TestDrive 'default-title.html'
        Export-DetectionResultsHtml -DataTable (New-Table @('t', 'File', 'Positive', 'a', 'b', 'c', 'Pass')) -OutputPath $out
        (Read-Html $out) | Should -Match '<h1>Detection Test Results</h1>'
    }

    It 'marks <Value> with <Glyph> and the class <Class>' -ForEach @(
        @{ Value = 'Pass';     Glyph = '&#x2713;'; Class = 'result-pass' }
        @{ Value = 'Detected'; Glyph = '&#x2713;'; Class = 'result-pass' }
        @{ Value = 'pass';     Glyph = '&#x2713;'; Class = 'result-pass' }
        @{ Value = 'Fail';     Glyph = '&#x2717;'; Class = 'result-fail' }
        @{ Value = 'Error';    Glyph = '&#x2717;'; Class = 'result-fail' }
        @{ Value = '';         Glyph = '&#x2717;'; Class = 'result-fail' }
    ) {
        $out = Join-Path $TestDrive 'glyph.html'
        Export-DetectionResultsHtml -DataTable (New-Table @('t', 'File', 'Positive', 'a', 'b', 'c', $Value)) -OutputPath $out
        (Read-Html $out) | Should -Match ("<td class='{0}'><span class='glyph'>{1}</span>{2}</td>" -f $Class, [regex]::Escape($Glyph), [regex]::Escape($Value))
    }

    It 'adds a glyph only in the Result column' {
        $out = Join-Path $TestDrive 'glyph-only-result.html'
        Export-DetectionResultsHtml -DataTable (New-Table @('t', 'File', 'Positive', 'Pass', 'Pass', 'Pass', 'Pass')) -OutputPath $out
        ([regex]::Matches((Read-Html $out), "class='glyph'")).Count | Should -Be 1
    }

    It 'adds no glyph when the table has no Result column' {
        $t = New-ResultTable -Column @('Target', 'Found')
        [void]$t.Rows.Add('a', 'Pass')
        $out = Join-Path $TestDrive 'no-result.html'
        Export-DetectionResultsHtml -DataTable $t -OutputPath $out
        Read-Html $out | Should -Not -Match "class='glyph'"
    }

    It 'encodes <Name> in a cell, a column name and the title' -ForEach @(
        @{ Name = 'angle brackets'; Raw = '<script>alert(1)</script>'; Encoded = '&lt;script&gt;alert(1)&lt;/script&gt;' }
        @{ Name = 'an ampersand';   Raw = 'a & b';                    Encoded = 'a &amp; b' }
        @{ Name = 'double quotes';  Raw = 'say "hi"';                 Encoded = 'say &quot;hi&quot;' }
        @{ Name = 'a single quote'; Raw = "it's";                     Encoded = 'it&#39;s' }
        @{ Name = 'an entity';      Raw = '&lt;';                     Encoded = '&amp;lt;' }
    ) {
        $t = New-ResultTable -Column @('Target', 'Result')
        [void]$t.Rows.Add($Raw, 'Pass')
        $t.Columns[0].ColumnName = 'Col ' + $Raw
        $out = Join-Path $TestDrive 'encode.html'
        Export-DetectionResultsHtml -DataTable $t -OutputPath $out -ReportTitle $Raw
        $html = Read-Html $out
        $html | Should -Match ('<td>' + [regex]::Escape($Encoded) + '</td>')
        $html | Should -Match ('<th>Col ' + [regex]::Escape($Encoded) + '</th>')
        $html | Should -Match ('<title>' + [regex]::Escape($Encoded) + '</title>')
        $html | Should -Not -Match '<script>'
    }

    It 'keeps non-ASCII text and line breaks in a cell' {
        # HtmlEncode writes the Latin-1 letters as numeric entities and keeps the other letters, so decode before the comparison.
        $text = "Caf$([char]0xE9) $([char]0x20AC) $([char]0x65E5)$([char]0x672C)`nsecond"
        $out = Join-Path $TestDrive 'unicode.html'
        Export-DetectionResultsHtml -DataTable (New-Table @('t', 'Script', 'Positive', 'x', 'y', $text, 'Pass')) -OutputPath $out
        $html = Read-Html $out
        $html | Should -BeLike "*$([char]0x20AC) $([char]0x65E5)$([char]0x672C)*"
        [System.Net.WebUtility]::HtmlDecode($html) | Should -BeLike "*<td>$text</td>*"
    }

    It 'writes a page for a table without rows' {
        $out = Join-Path $TestDrive 'norows.html'
        Export-DetectionResultsHtml -DataTable (New-ResultTable) -OutputPath $out
        $html = Read-Html $out
        $html | Should -Match 'Tests: 0'
        $html | Should -Match '<tbody></tbody>'
    }

    It 'writes an empty cell for a database null' {
        $t = New-ResultTable
        $row = $t.NewRow(); $row['Time'] = 't'
        [void]$t.Rows.Add($row)
        $out = Join-Path $TestDrive 'null.html'
        Export-DetectionResultsHtml -DataTable $t -OutputPath $out
        Read-Html $out | Should -Match '<td></td>'
    }

    It 'creates the folder and replaces an existing file' {
        $out = Join-Path $TestDrive 'html folder\deep\out.html'
        Export-DetectionResultsHtml -DataTable (New-Table @('t', 'File', 'Positive', 'a', 'b', 'c', 'Pass')) -OutputPath $out
        Test-Path -LiteralPath $out | Should -BeTrue
        Set-Content -LiteralPath $out -Value ('old ' * 500)
        Export-DetectionResultsHtml -DataTable (New-Table @('t', 'File', 'Positive', 'a', 'b', 'c', 'Pass')) -OutputPath $out
        Read-Html $out | Should -Not -Match 'old old'
    }

    It 'throws when the file cannot be written' {
        $f = Join-Path $TestDrive 'iamafile.txt'
        Set-Content -LiteralPath $f -Value 'x'
        { Export-DetectionResultsHtml -DataTable (New-Table @('t', 'File', 'Positive', 'a', 'b', 'c', 'Pass')) -OutputPath (Join-Path $f 'out.html') } | Should -Throw
    }

    It 'logs the export once with the path' {
        $out = Join-Path $TestDrive 'logged.html'
        Export-DetectionResultsHtml -DataTable (New-Table @('t', 'File', 'Positive', 'a', 'b', 'c', 'Pass')) -OutputPath $out
        Should -Invoke Write-Log -ModuleName DetectionTesterCommon -ParameterFilter { $Message -eq "Exported HTML to $out" } -Times 1 -Exactly
    }
}
