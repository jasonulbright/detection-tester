# Changelog

All notable changes to Detection Method Tester are documented in this
file.

## [1.0.0] - 2026-05-02

Detection Method Tester is a local GUI for testing MECM application
detection methods (RegistryKeyValue, RegistryKey, File, Script,
Compound) against the machine it's running on, without deploying the
app through MECM. Ships as a release zip; extract and run
`start-detectiontester.ps1`. No MSI, no code signing, no NuGet.

### Features

- **Detection Tester** module — authors the five MECM detection types
  and evaluates them locally: RegistryKeyValue (Exists / IsEquals /
  GreaterEquals / etc.), RegistryKey existence, File path + version
  comparison, Script (PowerShell body), and Compound (AND / OR
  connector across clauses). Dynamic input panels swap based on the
  selected type; result panel shows pass / fail and the actual values
  observed on the host.
- **Hive selector** on registry detection types — HKLM or HKCU. HKCU
  covers per-user installs (Adobe Reader DC, Chrome user-mode,
  OneDrive).
- **Explicit 32 / 64-bit view selector** on registry detection types —
  removes ambiguity around WOW6432Node-vs-native paths.
- **Negative-test mode** — flips pass / fail logic so "not detected"
  is a passing test. Useful for verifying that an uninstall removed
  the ARP entry the clause targets.
- **Installed Applications** module — reads HKLM (x64 + WOW6432Node)
  and HKCU ARP hives. Filter by name, scope by Machine / User. "Use
  for detection" populates the Detection Tester registry fields with
  hive, architecture, key path, and DisplayVersion in one click.
  Detail panel exposes DisplayName, Publisher, DisplayVersion,
  Architecture, Scope, RegistryKey, UninstallString,
  QuietUninstallString, InstallLocation, InstallDate. Copy-details to
  clipboard.
- **Test history** grid — every evaluated clause logged with time,
  type, mode (Positive / Negative), target, expected, found, and
  result. Export to CSV / HTML; the "Use selected" button repopulates
  the input panel from the selected row.
- **Import stage-manifest** — loads `stage-manifest.json` from
  Application Packager and populates the clause authoring fields,
  including compound detection (AND / OR).
- **Options dialog** — About panel (version, license, module version)
  and Logging panel (current log path, log folder, Open log folder /
  Open app data folder buttons).
- Dark / Light theme toggle on the sidebar; live swap, no restart.
  Per-app data (logs, prefs, window state) lives under
  `%LOCALAPPDATA%\DetectionTester\`.
- **Title-bar drag fallback** — native `WM_NCHITTEST` hook plus a
  managed `DragMove` for the main window and modal Options dialog,
  so the title bar drags reliably under any host.

### Stack

- PowerShell 5.1 + .NET Framework 4.7.2+
- WPF + MahApps.Metro (vendored DLLs in `Lib\`)
- ConfigurationManager PowerShell module is **not** required (detection
  evaluation is purely local; no MECM connection needed)
