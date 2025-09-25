# AD BitLocker Escrow Inventory (GUI)

A PowerShell script that scans Active Directory computers, enumerates **BitLocker recovery keys** (msFVE‑RecoveryInformation) stored in AD, and displays the results in a **GUI** with filtering and **CSV export**.

![PowerShell](https://img.shields.io/badge/PowerShell-5.1%20%2F%207%2B-blue) ![Platform](https://img.shields.io/badge/Platform-Windows%20only-lightgrey) ![Requires](https://img.shields.io/badge/Requires-AD%20RSAT-orange)

---

## Features

* Scan the entire domain or a specific OU
* **Excludes Windows Server** by default (`OperatingSystem -match 'Server'`)
* Counts **msFVE‑RecoveryInformation** objects per computer
* Uses the most recent `whenCreated` as a proxy for *Encryption/Rotation Date*
* Interactive **Windows Forms GUI** with live filtering and CSV export
* Console summary (with/without keys)

---

## Requirements

* **PowerShell 7+**
* **Active Directory Module for Windows PowerShell (RSAT)** installed
* Network access to a domain controller
* AD read permissions for:

  * `Computer` objects (`OperatingSystem`, `lastLogonDate`)
  * child objects of type `msFVE‑RecoveryInformation`

> Note: The GUI uses Windows Forms. PowerShell 7+ is supported on Windows; it is not supported on non‑Windows platforms for this script.

---

## Installation

1. Install RSAT including the **ActiveDirectory** PowerShell module.
2. Save the script as `BitLocker-AdEscrow-Inventory.ps1`.
3. Run from **PowerShell 7+** with sufficient AD rights.

The script aborts early if the AD module is unavailable.

---

## Usage

### Syntax

```powershell
.\nBitLocker-AdEscrow-Inventory.ps1 [
  -SearchBase <string> ]
  [-IncludeServers]
  [-MaxLastLogonAgeDays <int>]
  [-Verbose]
```

### Parameters

* `-SearchBase <string>`
  LDAP path (OU/DN), e.g. `OU=Clients,OU=CH,DC=contoso,DC=local`. If omitted, the entire domain is scanned.

* `-IncludeServers` *(Switch)*
  Include Windows Server systems (default: excluded).

* `-MaxLastLogonAgeDays <int>`
  Exclude stale computer accounts older than N days by `lastLogonDate` (default: `0`, no filter).

### Examples

```powershell
# Scan entire domain, only clients, exclude devices with last logon > 90 days
.\nBitLocker-AdEscrow-Inventory.ps1 -MaxLastLogonAgeDays 90

# Scan a specific OU
.
\BitLocker-AdEscrow-Inventory.ps1 -SearchBase "OU=Workstations,DC=contoso,DC=local"

# Include servers as well
.
\BitLocker-AdEscrow-Inventory.ps1 -IncludeServers
```

---

## Output & GUI

* **Console summary**

  ```
  Summary: 84 of 120 devices have BitLocker recovery keys in AD. (36 without)
  ```
* **GUI**

  * Search box (filters across `ComputerName`, `OperatingSystem`, `DistinguishedName`)
  * Read‑only data grid
  * **Export CSV…** button (respects current filter)

### Columns

* `ComputerName`
* `OperatingSystem`
* `LastLogonDate`
* `HasRecoveryKeyInAD` (boolean)
* `RecoveryKeyCountAD`
* `EncryptionDate` (latest `whenCreated` of recovery objects)
* `DistinguishedName`

> CSV export replaces commas with semicolons to avoid splitting embedded commas.

---

## How It Works

1. Enumerates `Computer` objects via `Get-ADComputer`, optionally restricted by `-SearchBase`.
2. Excludes servers unless `-IncludeServers` is specified.
3. Applies stale filter if `-MaxLastLogonAgeDays > 0`.
4. For each computer, queries child objects of type `msFVE‑RecoveryInformation`:

   * Counts recovery keys
   * Takes the most recent `whenCreated` as `EncryptionDate`
5. Builds a `DataTable`, binds to a Windows Forms grid, shows the GUI, enables CSV export.

---

## Security & Privacy

* **Read‑only**: the script does not modify AD.
* Recovery key metadata can be sensitive. Handle exported CSVs accordingly.
* In restrictive environments, read access to `msFVE‑RecoveryInformation`  require delegated permissions.

---

## Troubleshooting

* **“ActiveDirectory module not found”**
  Install RSAT and ensure the ActiveDirectory module loads (`Import-Module ActiveDirectory`).

* **GUI does not appear**
  Run on Windows with PowerShell 5.1 or 7+; ensure the process is not constrained by AppLocker/Constrained Language Mode.

* **No keys shown**
  Verify that BitLocker recovery key escrow to AD is enabled by policy and that the account has read access to the recovery objects.

---

## License

MIT License

---

## Acknowledgments

* Built for quick visibility into BitLocker recovery key escrow state directly from AD.
