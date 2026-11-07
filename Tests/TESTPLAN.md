# Detection Method Tester: test plan

Measured 2026-11-06 with `Invoke-Checks.ps1` (Pester 5.7.1, PSScriptAnalyzer 1.25.0). Counts are tests as run; each data row counts as one test. Coverage is the share of module commands that the tests run.

| | Before | After |
|---|---|---|
| Tests | 73 on Windows PowerShell 5.1 and on PowerShell 7 (66 module and script tests, 7 documentation tests) | 817 on Windows PowerShell 5.1 and on PowerShell 7 (none skipped) |
| Module coverage | 73.9% (424 of 574 commands) | 96.3% on Windows PowerShell 5.1 (601 of 624 commands), 96.0% on PowerShell 7 (599 of 624) |
| Lab steps | none | none (the tool does not write to ConfigMgr) |

Run: `.\Invoke-Checks.ps1` (lint, tests, coverage floor 80%). It runs in CI on both hosts (`.github/workflows/checks.yml`).

## What is covered

| Area | File | Contents |
|---|---|---|
| Module loading, parse checks | `Module.Helpers.Tests.ps1` | Parse of the script and the module, exported functions, SuiteCommon logging |
| Value comparison, file version cleaning, booleans | `Module.Compare.Tests.ps1` | Equality, integer, version and String kinds, boundaries, bad input, unknown operators, version resource forms |
| Hive and path helpers | `Module.Paths.Tests.ps1` | Every prefix form, hive-only paths, empty and hostile input |
| Registry detection | `Module.Detection.Registry.Tests.ps1` | A throwaway key `HKCU\Software\ZZTEST-DT-<guid>`: every value kind, operators, both registry views, missing key and value, wildcard names, long paths, handle cleanup |
| File detection | `Module.Detection.File.Tests.ps1` | Files in a test folder, a compiled program with a known file version, environment variables, the WOW64 folder rewrite, folders, bad paths |
| Script detection | `Module.Detection.Script.Tests.ps1` | Output rules, errors, languages, isolation, timeout and its limits |
| Compound detection | `Module.Detection.Compound.Tests.ps1` | And and Or, clause types, defaults, unknown connectors and clause types, a clause that cannot run |
| Manifest import | `Module.Manifest.Tests.ps1` | Field names, defaults, zero values, booleans, compound blocks, UTF-8, read errors, bad JSON |
| Installed applications | `Module.Apps.Tests.ps1` | Throwaway ARP entries under HKCU, sorting, scope and hive pairs, keys that the key test finds |
| Export | `Module.Export.Tests.ps1` | CSV and HTML: round trip, escaping, special characters, formula cells, empty tables, write failures, log lines |
| Script logic | `Shell.Logic.Tests.ps1` | Combo values, history table, preferences, window state, panel switch, gather, validate, run, import, repopulate, export, filter, details, module switch, closing, release version |
| Script calls | `Script.Contract.Tests.ps1` | Every call from the script and the module into an app function or SuiteCommon names only parameters that exist |
| Windows | `Window.Tests.ps1` | Main window, both module views (five input panels), the Options dialog: names, titles, no clipped text and nothing off screen at 1024x720 and 1440x900 in both themes, text and control contrast |
| Handlers on real controls | `Window.Tests.ps1` | The script's own script blocks run in a real window: clicks, background script test, history, installed applications, theme change |
| User guide | `Documentation.Tests.ps1` | Links, images and pages of the user guide |

## Defects found by the tests

Each defect has a test that failed on the old code and passes now. The old code ran 91 failing tests of the new suite on Windows PowerShell 5.1.

| # | Defect | Test | Fix |
|---|---|---|---|
| 1 | An expected value of 0 in a manifest or a compound clause became an empty text | `Module.Manifest`, `Module.Detection.Compound` | `Test-DTHasValue` |
| 2 | A file version such as `1, 0, 0, 1` became `1,0,0,1`, which does not parse | `Module.Compare` | `ConvertTo-DTCleanFileVersion` |
| 3 | A registry read that failed left the root key open | `Module.Detection.Registry` (syntax check) | `finally` blocks |
| 4 | A manifest was read without UTF-8 (garbled text on Windows PowerShell 5.1), and a read error showed as "No Detection block" | `Module.Manifest` | `-Encoding UTF8`, `-ErrorAction Stop` |
| 5 | A path that is only a hive name (`HKLM`) gave the hive but the key path `HKLM` | `Module.Paths` | Path pattern |
| 6 | The help of `Compare-DetectionValue` said that `String` compares as text, but the code compares by version | `Module.Compare` | Decision: version order; the help text says so |
| 7 | An integer comparison rounded `1.5` to 2 | `Module.Compare` | `Int64.TryParse` |
| 8 | A compound clause that could not run ended the whole test, and an unknown connector gave a silent "not detected" | `Module.Detection.Compound` | Per-clause `try`, message for the connector |
| 9 | `Is64Bit` as the text "false" read as true | `Module.Compare`, `Module.Manifest`, `Module.Detection.Compound` | `ConvertTo-DTBoolean` |
| 10 | A folder with the name of the file counted as the file | `Module.Detection.File` | `-PathType Leaf` |
| 11 | The timeout message named the requested limit, not the limit that was used | `Module.Detection.Script` | `$limitSeconds` |
| 12 | The scope filter of Installed Applications did nothing (`continue` inside a `switch`) | `Shell.Logic`, `Window` (handlers) | `if` tests |
| 13 | A failed HTML export was logged as done, and an export error ended the tool | `Module.Export`, `Shell.Logic` | `-ErrorAction Stop`, `try` in the handlers |
| 14 | A CSV cell that starts with `=`, `+`, `-` or `@` runs as a formula in a spreadsheet | `Module.Export` | `ConvertTo-DTCsvSafeValue` (a single quote in front) |
| 15 | The title of the HTML report was split over three lines (operator precedence) | `Module.Export` | Parentheses |
| 16 | The Options dialog title bar had 3.3:1 contrast in the light theme | `Window` | Same brushes as the main window |
| 17 | The Copy details button was outside the window at 1024 pixels width | `Window` | Flexible filter column |
| 18 | A test of the old suite failed on every checkout with Windows line endings (the pattern ended with `\n`) | `Shell.Logic` | The test reads the syntax tree, not the text |

Items that look like defects and are not changed:

- A multi-string registry value shows its lines joined by a space (`a b`), and a binary value shows its bytes in decimal (`1 2 255`). Neither shows `System.String[]` or `System.Byte[]`. The tests pin this format.
- A hive name other than HKLM or HKCU (for example `HKEY_USERS`) reads as LocalMachine. The tool supports two hives. The tests pin this.
- `System.Version` ranks `1.0` below `1.0.0`. The tests pin this.

## Not covered, and why

- The tool does not call a ConfigMgr cmdlet, so there is no cmdlet contract test, no start-up test without the ConfigMgr module and no lab run.
- The Open and Save dialogs of the operating system (manifest import, CSV and HTML export) block the thread. Their handlers run in `Shell.Logic.Tests.ps1` with a fake dialog.
- The Copy details button writes to the clipboard. The text it copies is a function that the tests run.
- The title bar drag fallback, the window state at start-up and the update notice run in the script outside a function. The functions behind the window state are tested.
- Contrast of disabled controls, the sidebar button border, the light text boxes, the accent buttons and the selected list item misses the standard (text 4.5:1, borders 3:1). The cause is the shared SuiteCommon palette and the MahApps styles. `Tests/KnownContrastDefects.ps1` lists each violation. The window test fails on a new one and on a listed one that is fixed. The fix belongs in `app-packager-suite`.
- Script detection runs in the same process. A script that calls `[Environment]::Exit` ends the tool. The help of `Test-ScriptDetection` states this limit.
- All window tests are in one file, because a second STA runspace in the same process hangs the theme manager.

## Window guard

Window tests open a window only when the environment variable SUITE_ALLOW_WINDOW_TESTS is 1, and never on the dev PC (computer name AYW). Everywhere else they report Skipped with the reason, before any runspace or window exists. CI sets the variable in both jobs. Invoke-Checks.ps1 prints one line when the window tests are skipped. It then does not enforce the coverage floor, because the window tests may be needed to reach it. The coverage number is informational in that case. WindowGuard.Tests.ps1 tests the guard and checks that every test file that opens a window uses it.
