# Troubleshooting

This page lists the most common problems of Detection Tester with their cause and their fix.

## What it does

The page helps you find the fix for a message on the result panel. [Error messages](reference/error-messages.md) lists all messages.

## When to use it

Use it when a test gives an answer that you do not expect.

## Requirements

None.

## Steps

1. Read the result and the detail line.
2. Find the symptom in the tables.
3. Do the fix.
4. Select **Test detection** again.

## Tests

| Message or symptom | Cause | Fix |
|---|---|---|
| **Fail** with the detail `MATCH`. | The clause is detected and **Negative test** is selected. | To test for absence, keep the check box and find out why the clause is detected. Clear it only to test for presence. See [Negative test](core/negative-test.md). |
| A key exists in Regedit and the clause fails. | The view or the hive is not the one that you looked at. | Select the other view, or HKCU. |
| A script test shows **Fail** and no output. | The script wrote no text to the output. | Write a value to the output when the application is installed. |
| A test passes here and fails in a deployment. | The deployment runs as System and the tool as your account. | See [Permissions](reference/permissions.md). |

## Import

| Message or symptom | Cause | Fix |
|---|---|---|
| Import failed. | The file has no `Detection` block, or it is not valid JSON. | Import the `stage-manifest.json` of a staged build. |
| The type is Compound and the fields are read-only. | A Compound manifest was imported. | Edit the manifest and import it again. |

## Window

| Message or symptom | Cause | Fix |
|---|---|---|
| The tool opens on Installed Applications. | The tool saved that page. | Select **Detection Tester**. |
| The window opens in the wrong place. | The saved window state holds an old position. | Close the tool. Delete `%LOCALAPPDATA%\DetectionTester\windowstate.json`. Start the tool. |

## Links

None.
