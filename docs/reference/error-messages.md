# Error messages

This page lists the messages that Detection Tester shows. Each row gives the message, where it appears, the cause and the fix.

## What it does

The page lists the texts of the result panel and of the log.

## When to use it

Use it when the result panel shows a message that you do not know.

## Requirements

None.

## Messages

### Result panel

| Message | Cause | Fix |
|---|---|---|
| Invalid clause: Key path is required. | **Key path** is empty. | Enter the key path. |
| Invalid clause: Compound clause has no children (import a manifest with Clauses). | The type is Compound and no manifest was imported. | Import a manifest. |
| Value '`<name>`' not found in key | The value does not exist in that hive, view and key. | Check the hive, the view and the name. |
| NO MATCH (`<operator>`): '`<found>`' vs '`<expected>`' | The value differs from the expected value. | Correct the expected value or the operator. |
| MATCH (`<operator>`): '`<found>`' vs '`<expected>`' | The clause is detected. With a negative test this gives **Fail**. | None. |
| Import failed: Failed to import: `<reason>` | The manifest could not be read. See [Import a manifest](../core/import-manifest.md). | Fix the file. |
| Running... | A script test runs. | Wait. The time limit is 30 seconds. |
| Error | A script test failed or ran out of time. | Read the detail line. |

### Log

| Message | Cause |
|---|---|
| Test: type=`<type>` mode=`<mode>` result=`<result>` target='`<target>`' | One line for each test. |
| Import-DetectionManifest failed: `<reason>` | A manifest import failed. |
| Imported manifest: type=`<type>` from `<path>` | A manifest import worked. |
| Repopulate: restored full clause from history (type=`<type>`, negative=`<n>`) | Use selected. |

## Steps

1. Find the message in the tables.
2. Do the fix.

## Limits

The tool shows the error text of the system as it is.

## Troubleshooting

See [Troubleshooting](../troubleshooting.md).

## Links

None.
