# Permissions

This page lists the rights that each clause needs.

## What it does

Every clause runs as the Windows account that started the tool. The tool never asks for elevation.

## When to use it

Use it to find out why a clause gives another result than the same clause on a device.

## Requirements

None.

## Rights by clause

| Clause | Right |
|---|---|
| RegistryKeyValue, RegistryKey | Read access to the key. Access to HKLM depends on the permissions of the key. HKCU is the hive of the account that started the tool. |
| File | Read access to the folder. |
| Script | The rights of the account. The script runs in a new PowerShell runspace on this computer. |
| Installed Applications | Read access to the three ARP keys. |

A Configuration Manager client runs a detection method as the System account for a machine deployment. A clause can pass here and fail there, or the other way. Start the tool as the same account to compare.

## Steps

### Test as another account

Start the tool with **Run as different user** or from a PowerShell prompt of that account.

## Limits

The tool does not switch accounts. To test as System, start the tool in a session that runs as System.

## Troubleshooting

See [Troubleshooting](../troubleshooting.md).

## Links

- [Create applications in Configuration Manager](https://learn.microsoft.com/intune/configmgr/apps/deploy-use/create-applications): Describes how a deployment type runs a detection method.
