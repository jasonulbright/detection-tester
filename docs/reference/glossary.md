# Glossary

This page defines the terms of Detection Tester.

## What it does

The page gives one meaning for each term.

## When to use it

Use it when a term is not clear.

## Requirements

None.

## Terms

| Term | Meaning |
|---|---|
| ARP | Add Remove Programs: the uninstall registry keys that list installed applications. |
| Clause | One detection rule: a registry value, a registry key, a file or a script. |
| Compound | A detection that combines clauses with And or Or. The tool tests it and cannot author it. |
| Connector | The word `And` or `Or` that combines the clauses of a Compound detection. |
| Detection method | The rule that tells Configuration Manager that an application is installed. |
| Hive | The root of a registry path: HKLM (the machine) or HKCU (the current user). |
| Negative test | A test that passes when the clause is not detected. |
| Stage manifest | The `stage-manifest.json` that Application Packager writes. Its `Detection` block holds the detection method. |
| View | The 64-bit registry view or the 32-bit view (WOW6432Node). |

## Steps

None.

## Limits

None.

## Troubleshooting

See [Troubleshooting](../troubleshooting.md).

## Links

None.
