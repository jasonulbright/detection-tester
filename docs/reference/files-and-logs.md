# Files and logs

This page lists every folder and file that Detection Tester writes.

## What it does

The tool writes in `%LOCALAPPDATA%\DetectionTester` and in the folder that you choose for an export. It writes nothing in the registry and nothing in the folder of the script.

## When to use it

Use it to find a log or an export, or to clean up.

## Requirements

None.

## Files

| Path | Content | Written when |
|---|---|---|
| `%LOCALAPPDATA%\DetectionTester\prefs.json` | The theme and the open page. | You change the theme or the page. |
| `%LOCALAPPDATA%\DetectionTester\windowstate.json` | The window size and place. | You close the window. |
| `%LOCALAPPDATA%\DetectionTester\logs\DetectionTester-<yyyyMMdd-HHmmss>.log` | The log of one session: each test, each import and each Use selected. | During the session. |
| `%LOCALAPPDATA%\DetectionTester\reports\detection-history-<date>-<time>.csv` or `.html` | An export of the history. The Save As dialog starts in this folder. | **Export CSV** or **Export HTML**. |

A log line is `[time] [level] text`. A test writes `Test: type=<type> mode=<mode> result=<result> target='<target>'`.

## Steps

### Clean up

1. Close the tool.
2. Delete the logs and the exports that you do not need.

## Limits

- The tool never deletes a log or an export.

## Troubleshooting

None.

## Links

None.
