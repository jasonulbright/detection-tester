# Capture the guide images

The images in `docs/images` show the main window and the Options dialog. `tools/guide-capture/Invoke-GuideCapture.ps1` takes all of them again with one command. The README shows `docs/images/start-first-run.png`.

## What the tool takes

The tool copies the app to a lab computer, starts it in an interactive session, drives it with UI Automation and copies the PNG files back. The steps run detection tests on the computer: they read the registry and the file system, and they run one PowerShell script. They register no detection method.

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

- The lab computer needs an interactive session of the account in `-LogonUser`.
- One capture runs at a time on a lab computer. The script takes the lock `C:\guide\capture.lock` and waits up to `-LockWaitMinutes` (default 90) for another capture.
- `-Section 'light theme'` runs one section. `-Image options-about` copies only that image into `docs/images`.
- Only images that already exist in `docs/images` are replaced. `-AddNew` also adds new ones.
- A run without `-Section` stops and replaces no image when a step did not take an image of `docs/images`.
- The steps reset the preferences of the lab account (`%LOCALAPPDATA%\DetectionTester\prefs.json`) before and after the run. They remove the history exports that the run wrote.
- The test `Tests\GuideCapture.Guard.Tests.ps1` fails when an image under `docs/images` has no step.

## One pass: check and capture

The capture run is also the check run. It starts the app maximized on a 1080p session (the work area is 1920x1032 at 100 percent), shoots every window in both themes, and samples each shot: the mean luminance, and a warning for a blank image. The run ends with one line, for example `Capture verdict: PASS. Images produced 38, probe samples 38, warnings 0, errors 0.` It writes `tools\guide-capture\produced-images.json`, the list of the images that the run produced. The test `Tests\GuideImages.Guard.Tests.ps1` fails when an image under `docs/images` is not in that list.

The window tests (`Tests\Density.Window.Tests.ps1` and the others) measure the layout at the work area of the session. Run them with `Tests\Run-VisibleProbe.ps1` in the same session when you change a size.
