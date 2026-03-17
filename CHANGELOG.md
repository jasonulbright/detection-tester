# Changelog

All notable changes to Detection Method Tester are documented in this file.

## [1.0.2] - 2026-03-17

### Added
- **App detail panel** on Installed Applications tab -- SplitContainer with grid on top, detail panel below showing all ARP fields for the selected application: DisplayName, Publisher, DisplayVersion, Architecture, RegistryKey, UninstallString, QuietUninstallString, InstallLocation, InstallDate
- **Copy Details** button copies all fields to clipboard in key-value format for easy paste into packager scripts or documentation
- `QuietUninstallString`, `InstallLocation`, and `InstallDate` now read from ARP registry during enumeration

## [1.0.1] - 2026-03-04

### Fixed
- Dark mode restart now captures script path at function scope (`$scriptFile = Join-Path $PSScriptRoot ...`) instead of using `$MyInvocation.ScriptName` which resolves to empty inside event handler scriptblocks

---

## [1.0.0] - 2026-03-05

### Added
- **Detection Tester tab** -- test MECM detection methods against the local machine without deploying through MECM; supports RegistryKeyValue, RegistryKey, File, and Script detection types
- **Installed Applications tab** -- enumerates both ARP registry hives (x64 + WOW6432Node) with text filter and "Use for Detection" button that populates the Detection Tester fields
- Dynamic input panels that swap based on selected detection type
- Color-coded results panel with detailed match/no-match output
- Test history grid with time, type, target, expected, found, and result columns
- Import manifest button for loading stage-manifest.json files from Application Packager
- Compound detection support (And/Or) via manifest import
- Export test history to CSV or HTML from history grid context menu
- Dark/light theme, window state persistence, preferences dialog
