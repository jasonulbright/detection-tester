# Capture the guide images

The images in `docs/images` show the main window and the Options dialog. `tools/guide-capture/Invoke-GuideCapture.ps1` takes some of them again with one command. The README shows `docs/images/start-first-run.png`.

## What the tool takes

The tool copies the app to a lab workstation, starts it in an interactive session, drives it with UI Automation and copies the PNG files back. The steps only read. They register no detection method.

| Section | Images |
|---|---|
| dark theme | `start-first-run`, `installed-panel`, `options-about`, `options-logging` |
| light theme | `light-detection`, `light-installed`, `light-options` |

The other images of the guide (the Registry, File, Script and Compound results, the history and the import pages) are not taken by the tool yet. Take them by hand, or add a section to `tools/guide-capture/steps-guide.ps1`.

## Run it

```powershell
.\tools\guide-capture\Invoke-GuideCapture.ps1 -CredentialScript <credential script> -Target <workstation> -LogonUser <domain\account>
```

- The workstation needs an interactive session of the account in `-LogonUser`.
- One capture runs at a time on a workstation. The script takes the lock `C:\guide\capture.lock` and waits up to `-LockWaitMinutes` (default 90) for another capture.
- `-Section 'light theme'` runs one section. `-Image options-about` copies only that image into `docs/images`.
- Only images that already exist in `docs/images` are replaced. `-AddNew` also adds new ones.
- The steps reset the preferences of the lab account (`%LOCALAPPDATA%\DetectionTester\prefs.json`) before and after the run.
