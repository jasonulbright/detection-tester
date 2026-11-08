# Releasing Detection Method Tester

## 1. Test

Run the checks under Windows PowerShell 5.1 and under PowerShell 7. The command runs the linter, all tests and the coverage floor. The release stops if either run fails.

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File .\Invoke-Checks.ps1
pwsh -NoProfile -File .\Invoke-Checks.ps1
```

The same command runs in CI on both hosts (`.github/workflows/checks.yml`). The window tests open windows off screen and need a desktop session.

## 2. Version

The version is `YYYY.MM.DD.BBBB`: the release date, then a four-digit, zero-padded build number. The build number increases by 1 for each release and never resets. Build 0007 is the first release with this scheme; the 6 releases before it used `1.x.y` numbers. Keep the zero-padded text everywhere; `[version]` drops the leading zeros.

Set the same version in three places:

- `Module/DetectionTesterCommon.psd1`, `ModuleVersion`
- `start-detectiontester.ps1`, header line `Version    : <ver>`
- `CHANGELOG.md`, the top heading `## [<ver>] - <date>`

Add the new `CHANGELOG.md` entry at the top. Do not edit the entries of earlier releases.

## 3. Shared module

Check the vendored SuiteCommon copy for drift. Sync it if the check reports drift.

```powershell
C:\projects\app-packager-suite\sync-suitecommon.ps1 -Consumer C:\projects\detection-tester -Check
```

## 4. Commit, tag, and package

Commit to `master`. Every commit to `master` is part of a release: the suite installer build refuses a component whose `master` is ahead of its latest tag. Tag the commit `v<ver>` with an annotated tag. Build the zip from the tag. The zip excludes the tests, the check script, the CI workflow and this file.

```bash
git tag -a v<ver> -m v<ver>
git archive --format=zip -o DetectionMethodTester-<ver>.zip v<ver> -- . ':(exclude)Tests' ':(exclude)*.Tests.ps1' ':(exclude)Invoke-Checks.ps1' ':(exclude).github' ':(exclude)RELEASING.md'
sha256sum DetectionMethodTester-<ver>.zip | sed 's/ \*/  /' > checksums.txt
```

Extract the zip to a temporary folder. Import `Module/DetectionTesterCommon.psd1` under Windows PowerShell 5.1. The import must succeed.

## 5. Publish

Push `master` and the tag. Create the GitHub release with the title `v<ver>` and two assets: the zip and `checksums.txt`. The release must not be a draft. The release notes have a `##` headline with one concrete outcome metric, the `###` sections of the changelog entry, and the footer `Full changelog: CHANGELOG.md`.
