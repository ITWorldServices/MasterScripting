# Single-Server Discovery — Phase 1

`Invoke-ServerDiscovery.ps1` creates a timestamped copy of the TPT Server Migration Planning workbook and inventories only the Windows Server on which it is executed.

## Safety boundaries

The Phase 1 script:

- Does not enumerate computer objects from Active Directory.
- Does not use PowerShell remoting or WinRM.
- Installs only the pinned ImportExcel 7.8.10 module when it is missing; no other application software is installed.
- Does not use `Win32_Product`.
- Does not recurse through file shares or calculate folder sizes.
- Reads share-root NTFS permissions only.
- Continues after an individual collector fails.
- Records collector status and errors in the `Diagnostics` worksheet.
- Copies the workbook template; it does not overwrite the repository template.

Run the first test during a normal maintenance window. The collection is read-only, but roles such as DHCP, DNS, and AD DS can contain a large amount of configuration data.

## Requirements

- Windows Server 2016 or later
- Windows PowerShell 5.1 or PowerShell 7 on Windows
- Local administrator is recommended
- Outbound HTTPS access to PowerShell Gallery and NuGet endpoints for the first run
- The workbook template `TPT - Server Migration Planning Document.xlsx`
- No Git installation, clone, or GitHub connection is required on the server

ImportExcel 7.8.10 is installed automatically for the current user when missing. The bootstrap:

1. Enables TLS 1.2 for the current PowerShell process.
2. Installs the NuGet package provider when missing.
3. Registers the default PSGallery repository when missing.
4. Temporarily marks PSGallery trusted to prevent an interactive prompt.
5. Installs the pinned ImportExcel version.
6. Restores the repository's previous trust policy.

This requires outbound HTTPS access. It does not connect to GitHub.

## First test

1. Download these two files from GitHub on your administrative workstation:
   - `Invoke-ServerDiscovery.ps1`
   - `TPT - Server Migration Planning Document.xlsx`
2. Transfer both files to `C:\Temp` on the test server using your approved method.
3. Open an elevated PowerShell session on the server.
4. Run:

```powershell
Set-Location C:\Temp
Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass
.\Invoke-ServerDiscovery.ps1
```

The script finds the workbook beside itself and writes the result directly to `C:\Temp`:

```text
C:\Temp\SERVERNAME-ServerDiscovery-yyyyMMdd-HHmmss.xlsx
```

No Git client, repository clone, GitHub sign-in, or outbound GitHub connection is used by the script. On the first run, it connects to PowerShell Gallery/NuGet only when ImportExcel 7.8.10 is not already installed.

A different template or output location can still be supplied explicitly:

```powershell
.\Invoke-ServerDiscovery.ps1 -TemplatePath D:\Templates\Planning.xlsx -OutputDirectory D:\Discovery
```

## What Phase 1 collects

| Area | Source | Workbook output |
|---|---|---|
| Hardware and OS | CIM: ComputerSystem, OperatingSystem, Processor, BIOS | Server Info; System Inventory |
| IPv4 network configuration | Get-NetIPConfiguration | Server Info; Network Inventory |
| Fixed disks | Win32_LogicalDisk, DriveType 3 | Server Info; Disk Inventory |
| Listening TCP ports | Get-NetTCPConnection | Server Info; Listening Ports |
| Installed applications | 32-bit and 64-bit HKLM uninstall registry | LoB Applications; Application Inventory |
| Services and service identities | Win32_Service | Services |
| Installed roles and features | Get-WindowsFeature | Server Roles; Server Features; Installed Features |
| Azure agents and applications | Application/service name matching | Azure Agents-Apps-Services |
| SMB shares and root permissions | Get-SmbShare, Get-SmbShareAccess, Get-Acl | File Server Shares; File Server Security; SMB Share Inventory |
| Printers | Get-Printer | Printer Inventory |
| Local users and groups | Microsoft.PowerShell.LocalAccounts | Local Accounts; Local Group Membership |
| Scheduled tasks | Get-ScheduledTask | Scheduled Tasks; Task Inventory |
| Local DC facts and FSMO roles | ActiveDirectory module, only when AD DS is installed | Active Directory; Domain Controller Inventory |
| Local DHCP scopes | DHCP cmdlets, only when DHCP is installed | DHCP; DHCP Scope Inventory |
| Local DNS zones | DNS cmdlets, only when DNS is installed | DNS Zone Inventory |
| Collector results | Internal error handling | Diagnostics |

## Intentionally manual or deferred

These template fields cannot be discovered reliably from one server without customer-specific systems or decisions:

- Notes, migration disposition, purpose, location, INC status/site
- Backup product status and policy compliance
- Warranty and support entitlement
- Application installation media, migration instructions, licensing, contracts, vendor support, and documentation
- Testing and QA decisions
- Risk Management decisions
- Full AD user and security-group export
- Environment-wide server enumeration
- Recursive file/folder size and ACL analysis
- DNS record export and replication health
- DHCP exclusions, reservations, options, NAP, and failover details
- Print-share security details

These are candidates for Phase 2 collectors after the single-server output is reviewed.

## Reviewing the result

Before expanding to all servers, verify:

1. The copied workbook opens without a repair warning.
2. Existing manual/planning worksheets and formatting remain present.
3. `Server Info` values match the test server.
4. Role and feature Yes/No results match Server Manager.
5. Applications are complete enough without `Win32_Product`.
6. Share permissions represent only share roots, as intended.
7. Failed or partial collectors are explained on `Diagnostics`.
8. Sensitive workbook output is stored in an approved location and is not committed to Git.

## Existing-script decisions

The earlier discovery scripts remain unchanged for comparison. Phase 1 consolidates their useful intent but replaces these patterns:

- AD-wide discovery and repeated remote sessions are deferred.
- `Win32_Product` application inventory is replaced with uninstall-registry enumeration.
- The missing ImportExcel dependency is bootstrapped from PSGallery at the pinned version 7.8.10.
- Wide application columns are retained for compatibility, with a normalized Application Inventory sheet added.
- Collector-level error handling replaces all-or-nothing execution.
- `Export-Excel -Show` is not used, allowing headless execution.

## Next phase

After one test workbook is reviewed, Phase 2 can add controlled AD discovery, include/exclude filters, OU scoping, remoting timeouts, concurrency limits, offline-server reporting, and consolidated multi-server output.
