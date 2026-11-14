# Changelog

## [Unreleased]

### Fixed

- Show button borders at 1 px in gray, with no heavy border in the text color.
- Add guards for button styles, muted text brushes and the dialog theme call.
- Make the PowerShell sources and the files they read plain ASCII, and add the encoding guard.

## [2026.11.14.0018] - 2026-11-14

## 4.5:1 contrast for the inactive title bar in both themes

### Fixed

- Show inactive title bar text in white at 4.5:1 contrast or more in both themes.
- Use shared border and disabled-text styles for text boxes, combo boxes and grid headers.
- Show test windows at 80,80 when SUITE_VISIBLE_WINDOWS is 1, for visible runs.

## [Unreleased]

### Fixed

- Show button borders at 1 px in gray, with no heavy border in the text color.
- Add guards for button styles, muted text brushes and the dialog theme call.
- Make the PowerShell sources and the files they read plain ASCII, and add the encoding guard.

## [2026.11.13.0017] - 2026-11-13

## 1 shared footer style and 2 new guard test files

### Fixed

- Restore the weight of the page title, the section headers, the log header, the history header, the result badge, the About keys and the details keys to SemiBold. The light theme change had made them bold. Only field titles and the app title are bold.

### Added

- Add FontWeight.Guard.Tests.ps1 and FontWeightGuardHelpers.ps1. They fail when a bold weight is set on a label that is not a field title or an app title, in XAML, in the XAML of the scripts and in code that sets a weight at run time.
- Add Footer.Guard.Tests.ps1. It checks the footer status element, the shared footer font and the footer calls. It also checks that the tool has no site code box and no SMS Provider box. It has bad-sample tests.
- Add a bad-sample test to FontWeight.Guard.Tests.ps1: a bold label before a panel that holds an input is not a field title.

### Changed

- Show the status text in the shared footer style of SuiteCommon: the same font in both themes, and no font setting on the status element. The tool has no Configuration Manager connection, so the footer shows no site text.
- Refresh the vendored SuiteCommon library to release 2026.11.13.0082.
- Tighten the field title rule of FontWeightGuardHelpers.ps1. A label counts as a field title only when an input follows or precedes it directly, or shares its grid row.
- Update the README text. Retake the images of the user guide and of the README.

## [Unreleased]

### Fixed

- Show button borders at 1 px in gray, with no heavy border in the text color.
- Add guards for button styles, muted text brushes and the dialog theme call.
- Make the PowerShell sources and the files they read plain ASCII, and add the encoding guard.

## [2026.11.12.0016] - 2026-11-12

## 1 shared themed style for every button in both themes

### Fixed

- Show white text on blue buttons and selected items in the light theme.
- Keep text contrast on buttons when the pointer is over them or they are pressed.
- Raise the dark theme button border to a contrast of 3:1.
- Show a fixed message in the window when an error occurs, not the raw exception text.

### Changed

- Use one themed style for every button.
- Use one gray style for secondary dialog buttons such as Cancel.
- Draw selected rows, cells and list items in the theme highlight color.
- Write titles, buttons and field labels in Title Case. Show titles and field labels in bold.
- Update the shared SuiteCommon module to 2026.11.12.0081.
- Retake the images of the user guide.
- Add tests that check the theme rules.

## [Unreleased]

### Fixed

- Show button borders at 1 px in gray, with no heavy border in the text color.
- Add guards for button styles, muted text brushes and the dialog theme call.
- Make the PowerShell sources and the files they read plain ASCII, and add the encoding guard.

## [2026.11.07.0015] - 2026-11-07

## 21 defects fixed that the new test suite found

### Fixed

- Compare an expected value of 0 as the number 0 in manifests and compound tests.
- Read file versions such as "1, 0, 0, 1" as versions.
- Read manifests with non-ASCII text correctly in Windows PowerShell 5.1.
- Show a read error for a manifest that cannot be read, not "No Detection block".
- Treat a path that is only a hive name as an empty key path.
- Stop integer comparisons from rounding values such as 1.5.
- Fail only the clause that cannot run in a compound test.
- Name an unsupported connector in the compound result.
- Read the text "false" in Is64Bit as false.
- Stop a folder with the name of the file from counting as the file.
- Report the real limit in the script timeout message.
- Close registry keys when a read fails.
- Apply the Machine and User scope filter in Installed Applications.
- Show an export error in the result panel, and stop reporting a failed HTML export as done.
- Prefix CSV cells that start with = + - or @ with a quote, so a spreadsheet does not run them.
- Write the HTML report title on one line.
- Raise the contrast of the Options title bar in the light theme.
- Keep the Copy details button inside the window at 1024 pixels width.
- Use the default for an Is64Bit value that is not a boolean, text or number.
- Prefix CSV column names that start with = + - or @ with a quote.
- Keep the value of every column in the CSV export when two column names are the same after the quote prefix.

### Changed

- Update the shared SuiteCommon module to 2026.11.07.0075.
- Run window tests only when the SUITE_ALLOW_WINDOW_TESTS opt-in is set.

## [Unreleased]

### Fixed

- Show button borders at 1 px in gray, with no heavy border in the text color.
- Add guards for button styles, muted text brushes and the dialog theme call.
- Make the PowerShell sources and the files they read plain ASCII, and add the encoding guard.

## [2026.11.07.0014] - 2026-11-07

## 1 user guide ships with the release

### Added

- Add the user guide (docs/index.md) with images of the running window.

## [Unreleased]

### Fixed

- Show button borders at 1 px in gray, with no heavy border in the text color.
- Add guards for button styles, muted text brushes and the dialog theme call.
- Make the PowerShell sources and the files they read plain ASCII, and add the encoding guard.

## [2026.11.06.0013] - 2026-11-06

## 2 crashes, 1 clipped header and 1 stale result fixed that the user guide found

### Fixed

- Stop a crash after a Compound import. A second import of a manifest, or Use selected on a Compound history row, ended the tool. The window state kept the clause data under a name that differs from a control name only in case.
- Stop a crash at Export CSV and Export HTML. The history table was returned as its rows, not as one table.
- Widen the Architecture column of Installed Applications so that its header is not cut.
- Clear the old result when a clause is filled from an installed application or from a manifest. The result of an earlier failed import stayed on the page.

## [Unreleased]

### Fixed

- Show button borders at 1 px in gray, with no heavy border in the text color.
- Add guards for button styles, muted text brushes and the dialog theme call.
- Make the PowerShell sources and the files they read plain ASCII, and add the encoding guard.

## [2026.10.12.0012] - 2026-10-12

## 1 PowerShell 7 fix

### Fixed

- Keep the PowerShell 7 module roots when the tool runs under PowerShell 7.

## [Unreleased]

### Fixed

- Show button borders at 1 px in gray, with no heavy border in the text color.
- Add guards for button styles, muted text brushes and the dialog theme call.
- Make the PowerShell sources and the files they read plain ASCII, and add the encoding guard.

## [2026.10.11.0011] - 2026-10-11

## 1 window fix

### Fixed

- Double-click the title bar to maximize or restore the window.

## [Unreleased]

### Fixed

- Show button borders at 1 px in gray, with no heavy border in the text color.
- Add guards for button styles, muted text brushes and the dialog theme call.
- Make the PowerShell sources and the files they read plain ASCII, and add the encoding guard.

## [2026.10.11.0010] - 2026-10-11

## 1 suite update notice

### Added

- Show a notice when a newer suite release is available; the launcher installs it.

All notable changes to Detection Method Tester are documented in this
file.

## [Unreleased]

### Fixed

- Show button borders at 1 px in gray, with no heavy border in the text color.
- Add guards for button styles, muted text brushes and the dialog theme call.
- Make the PowerShell sources and the files they read plain ASCII, and add the encoding guard.

## [2026.10.07.0009] - 2026-10-07

### Changed

- Sync the shared SuiteCommon module to 2026.10.07.0049.

## [Unreleased]

### Fixed

- Show button borders at 1 px in gray, with no heavy border in the text color.
- Add guards for button styles, muted text brushes and the dialog theme call.
- Make the PowerShell sources and the files they read plain ASCII, and add the encoding guard.

## [2026.09.21.0008] - 2026-09-21

### Changed

- Update the README screenshot.

## [Unreleased]

### Fixed

- Show button borders at 1 px in gray, with no heavy border in the text color.
- Add guards for button styles, muted text brushes and the dialog theme call.
- Make the PowerShell sources and the files they read plain ASCII, and add the encoding guard.

## [2026.09.21.0007] - 2026-09-21

### Changed

- Use the product name Configuration Manager in the application and the documentation.
- Use date versions.
- Update the shared SuiteCommon module to 2026.09.21.0031.

## [1.2.3] - 2026-09-04

### Fixed

- **Version labels read the script header.** The sidebar version and the
  About panel carried literal version strings that no release updated;
  both now read the entry script's `Version` header at startup, so the
  window always names the version that is actually installed.

## [1.2.2] - 2026-09-04

### Changed

- **Vendored `SuiteCommon` 0.4.3.** The module repairs the process
  PSModulePath at import: a Windows PowerShell process launched from
  PowerShell 7 inherits the 7.x module directories, and the background
  runspace opened later autoloaded a Microsoft.PowerShell.Utility without
  Get-FileHash or ConvertFrom-Json, so background operations failed with
  an unrelated "term not recognized". A background runspace whose module
  import fails is disposed and the original error thrown instead of being
  returned as an opened but unusable worker.

## [1.2.1] - 2026-08-16

### Changed

- **Vendored `SuiteCommon` 0.3.2.** Window restore applies the saved
  geometry before maximizing, so un-maximizing returns to the saved size
  instead of the XAML defaults.

## [1.2.0] - 2026-08-16

### Changed

- **The title-bar drag block now comes from the vendored `SuiteCommon`
  module** (0.3.0). Behavior gain: hook state no longer leaks when a
  window closes (removal previously keyed on the already-destroyed
  window handle).

## [1.1.0] - 2026-08-14

### Changed

- **Shared plumbing moved to the vendored `SuiteCommon` module.** Logging
  (`Initialize-Logging`, `Write-Log`) and preference persistence now load
  from `Lib\SuiteCommon\`, shared across the tool suite and synced from
  the suite-core repository instead of hand-edited per repo. Same log
  format; `Initialize-Logging` additionally gains `-Attach` and a DEBUG
  level exists behind the `SUITE_VERBOSE` environment variable.

## [1.0.0] - 2026-05-02

Detection Method Tester is a local GUI for testing MECM application
detection methods (RegistryKeyValue, RegistryKey, File, Script,
Compound) against the machine it's running on, without deploying the
app through MECM. Ships as a release zip; extract and run
`start-detectiontester.ps1`. No MSI, no code signing, no NuGet.

### Features

- **Detection Tester** module - authors the five MECM detection types
  and evaluates them locally: RegistryKeyValue (Exists / IsEquals /
  GreaterEquals / etc.), RegistryKey existence, File path + version
  comparison, Script (PowerShell body), and Compound (AND / OR
  connector across clauses). Dynamic input panels swap based on the
  selected type; result panel shows pass / fail and the actual values
  observed on the host.
- **Hive selector** on registry detection types - HKLM or HKCU. HKCU
  covers per-user installs (Adobe Reader DC, Chrome user-mode,
  OneDrive).
- **Explicit 32 / 64-bit view selector** on registry detection types -
  removes ambiguity around WOW6432Node-vs-native paths.
- **Negative-test mode** - flips pass / fail logic so "not detected"
  is a passing test. Useful for verifying that an uninstall removed
  the ARP entry the clause targets.
- **Installed Applications** module - reads HKLM (x64 + WOW6432Node)
  and HKCU ARP hives. Filter by name, scope by Machine / User. "Use
  for detection" populates the Detection Tester registry fields with
  hive, architecture, key path, and DisplayVersion in one click.
  Detail panel exposes DisplayName, Publisher, DisplayVersion,
  Architecture, Scope, RegistryKey, UninstallString,
  QuietUninstallString, InstallLocation, InstallDate. Copy-details to
  clipboard.
- **Test history** grid - every evaluated clause logged with time,
  type, mode (Positive / Negative), target, expected, found, and
  result. Export to CSV / HTML; the "Use selected" button repopulates
  the input panel from the selected row.
- **Import stage-manifest** - loads `stage-manifest.json` from
  Application Packager and populates the clause authoring fields,
  including compound detection (AND / OR).
- **Options dialog** - About panel (version, license, module version)
  and Logging panel (current log path, log folder, Open log folder /
  Open app data folder buttons).
- Dark / Light theme toggle on the sidebar; live swap, no restart.
  Per-app data (logs, prefs, window state) lives under
  `%LOCALAPPDATA%\DetectionTester\`.
- **Title-bar drag fallback** - native `WM_NCHITTEST` hook plus a
  managed `DragMove` for the main window and modal Options dialog,
  so the title bar drags reliably under any host.

### Stack

- PowerShell 5.1 + .NET Framework 4.7.2+
- WPF + MahApps.Metro (vendored DLLs in `Lib\`)
- ConfigurationManager PowerShell module is **not** required (detection
  evaluation is purely local; no MECM connection needed)
