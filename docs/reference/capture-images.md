# Capture the guide images

The images in `docs/images` show the main window and the Options dialog. `tools/guide-capture/Invoke-GuideCapture.ps1` takes all of them again with one command. The README shows `docs/images/start-first-run.png`.

## What the tool takes

The tool copies the app to a lab workstation, starts it in an interactive session, drives it with UI Automation and copies the PNG files back. The steps run detection tests on the workstation: they read the registry and the file system, and they run one PowerShell script. They register no detection method.

| Section | Images |
|---|---|
| first run | `start-first-run` |
| detection tests | `type-list`, `registry-value-pass`, `operator-list`, `registry-value-fail`, `registry-value-missing`, `registry-value-hkcu-32`, `type-registry-key`, `registry-key-pass`, `registry-key-fail`, `type-file`, `file-property-list`, `file-existence-pass`, `file-version-pass`, `file-fail`, `type-script`, `script-pass`, `script-fail`, `test-empty` |
| history and import | `import-compound`, `import-compound-run`, `history-filled`, `import-failed`, `use-selected`, `history-after-export` |
| negative test | `negative-test-fail`, `negative-test-pass` |
| installed applications | `installed-filter`, `installed-scope-list`, `installed-panel`, `installed-details`, `use-for-detection` |
| options | `options-about`, `options-logging` |
| restart | `start-reopen` |
| light theme | `light-detection`, `light-installed`, `light-options` |

A section uses the state that the sections before it leave: the history rows and the page. Select every section that an image needs, in the order of the table.

## Run it

```powershell
.\tools\guide-capture\Invoke-GuideCapture.ps1 -CredentialScript <credential script> -Target <workstation> -LogonUser <domain\account>
```

- The workstation needs an interactive session of the account in `-LogonUser`.
- One capture runs at a time on a workstation. The script takes the lock `C:\guide\capture.lock` and waits up to `-LockWaitMinutes` (default 90) for another capture.
- `-Section 'light theme'` runs one section. `-Image options-about` copies only that image into `docs/images`.
- Only images that already exist in `docs/images` are replaced. `-AddNew` also adds new ones.
- A run without `-Section` stops and replaces no image when a step did not take an image of `docs/images`.
- The steps reset the preferences of the lab account (`%LOCALAPPDATA%\DetectionTester\prefs.json`) before and after the run. They remove the history exports that the run wrote.
- The test `Tests\GuideCapture.Guard.Tests.ps1` fails when an image under `docs/images` has no step.
