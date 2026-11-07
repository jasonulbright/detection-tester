# Settings

This page lists what Detection Tester saves and where.

## What it does

The tool saves two small JSON files in `%LOCALAPPDATA%\DetectionTester`. It has no settings page.

## When to use it

Use it to reset the tool or to find where a choice is kept.

## Requirements

Write access to `%LOCALAPPDATA%\DetectionTester`.

## Settings

| Setting | Where in the window | File and key | Default | Saved |
|---|---|---|---|---|
| Theme | Toggle at the bottom of the sidebar | `prefs.json`, `DarkMode` | `true` (dark) | When you change it |
| Open page | The sidebar | `prefs.json`, `ActiveModule` (`DetectionTester` or `InstalledApps`) | The Detection Tester page | When you change the page |
| Window size and place | The window | `windowstate.json`: `Left`, `Top`, `Width`, `Height`, `Maximized` | 1100 by 750 in the middle of the screen | When you close the window |

The tool does not save the clause, the history or the filter.

## Steps

### Reset the tool

1. Close the tool.
2. Delete `prefs.json` and `windowstate.json` in `%LOCALAPPDATA%\DetectionTester`.
3. Start the tool. It opens with the defaults.

## Limits

- A window that was off the screen is moved to the middle at start. The tool compares the saved place with the bounding rectangle of all monitors. A window in a gap between monitors can stay off the screen. Reset the tool as described above.

## Troubleshooting

| Message or symptom | Cause | Fix |
|---|---|---|
| The tool opens on the Installed Applications page. | The tool saved that page. | Select **Detection Tester**, or reset the tool. |

## Links

None.
