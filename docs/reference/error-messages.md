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
| Invalid clause: Key path is required. | **Key Path** is empty. | Enter the key path. |
| Invalid clause: Compound clause has no children (import a manifest with Clauses). | The type is Compound and no manifest was imported. | Import a manifest. |
| Value '`<name>`' not found in key | The value does not exist in that hive, view and key. | Check the hive, the view and the name. |
| NO MATCH (`<operator>`): '`<found>`' vs '`<expected>`' | The value differs from the expected value. | Correct the expected value or the operator. |
| MATCH (`<operator>`): '`<found>`' vs '`<expected>`' | The clause is detected. With a negative test this gives **Fail**. | None. |
| Import failed: The manifest could not be imported. Select a stage-manifest.json file with a Detection block. | The manifest could not be read. The log has the reason. See [Import a manifest](../core/import-manifest.md). | Fix the file. |
| Export failed: The history could not be saved. Select a folder that you can write to. | The export could not write the file. The log has the reason. | Select another folder. |
| Running... | A script test runs. | Wait. The time limit is 30 seconds. |
| Error | A script test failed. | Read the detail line and the log. |

### Log

| Message | Cause |
|---|---|
| Test: type=`<type>` mode=`<mode>` result=`<result>` target='`<target>`' | One line for each test. |
| Import-DetectionManifest failed: `<reason>` | A manifest import failed. |
| Imported manifest: type=`<type>` from `<path>` | A manifest import worked. |
| Repopulate: restored full clause from history (type=`<type>`, negative=`<n>`) | Use Selected. |

## Steps

1. Find the message in the tables.
2. Do the fix.

## Limits

The tool shows the error text of the system as it is.

## Troubleshooting

See [Troubleshooting](../troubleshooting.md).

## Links

None.
