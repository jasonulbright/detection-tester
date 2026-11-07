# Network

This page lists every outbound call of Detection Tester.

## What it does

The tool tests this computer. It connects to no site. A copy that the suite installer installed can check GitHub for updates. A script that you test can make network calls.

## When to use it

Use it to set a firewall rule or to check what the tool sends.

## Requirements

None.

## Calls

| Call | Target | When | Content |
|---|---|---|---|
| `https://api.github.com/repos/<repository>/releases/latest` | GitHub | At start, about once a day, only when the tool is installed by the suite installer | The check for a newer suite version. The result goes to the status bar. |

A plain extract of the zip (no `suite-manifest.json`) makes no call. A failed lookup does not update the cache, so a later start can ask GitHub again within the day. Two starts at the same time can both ask before either saves the cache. A script clause can make any call that its script makes.

## Steps

None.

## Limits

The tool does not change a proxy setting. A failed update check is ignored.

## Troubleshooting

None.

## Links

None.
