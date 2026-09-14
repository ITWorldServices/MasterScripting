# Single-Server Discovery — Phase 1

`Invoke-ServerDiscovery.ps1` inventories the Windows Server on which it is executed and creates a clean, timestamped Excel workbook in `C:\Temp`.

The workbook is generated from collected data. The original migration-planning workbook is no longer required.

## Safety boundaries

The Phase 1 script:

- Does not enumerate computer objects from Active Directory.
- Does not use PowerShell remoting or WinRM.
- Installs only the pinned ImportExcel 7.8.10 module when it is missing.
- Does not use `Win32_Product`.
- Does not recurse through file shares or calculate folder sizes.
- Reads share-root NTFS permissions only.
- Continues after an individual collector fails.
- Records collector status, row counts, warnings, and errors in `Diagnostics`.
- Does not connect to GitHub.

Run the first test during a normal maintenance window. Collection is read-only, but AD DS, DHCP, and DNS can contain a large amount of configuration data.

## Requirements

- Windows Server 2016 or later
- Windows PowerShell 5.1 or PowerShell 7 on Windows
- Local administrator is recommended
- Outbound HTTPS access to PowerShell Gallery and NuGet endpoints for the first run
- No Git installation, repository clone, or GitHub connection on the server

ImportExcel 7.8.10 is installed automatically for the current user when missing. The bootstrap:

1. Enables TLS 1.2 for the current PowerShell process.
2. Installs the NuGet package provider when missing.
3. Registers the default PSGallery repository when missing.
4. Temporarily marks PSGallery trusted to prevent an interactive prompt.
5. Installs the pinned ImportExcel version.
6. Restores the repository's previous trust policy.

## Run the single-server test

1. Download `Invoke-ServerDiscovery.ps1` on your administrative workstation.
2. Transfer it to `C:\Scripts` or another approved directory on the test server.
3. Open an elevated PowerShell session.
4. Run:

```powershell
Set-Location C:\Scripts
Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass
.\Invoke-ServerDiscovery.ps1
```

The output is written directly to:

```text
C:\Temp\SERVERNAME-ServerDiscovery-yyyyMMdd-HHmmss.xlsx
```

Use a different output directory when required:

```powershell
.\Invoke-ServerDiscovery.ps1 -OutputDirectory D:\Discovery
```

### Optional legacy template parameter

`-TemplatePath` remains available for compatibility. When supplied, the script uses the file as the package starting point, then intentionally replaces its worksheets with the normalized discovery layout. Existing template tabs, placeholder values, and formatting are not retained.

```powershell
.\Invoke-ServerDiscovery.ps1 -TemplatePath D:\Templates\Planning.xlsx
```

## Workbook design

The workbook preserves the required discovery data without duplicating it across overlapping template sheets.

| Worksheet | Purpose |
|---|---|
| Discovery Summary | One-server overview, collection status, inventory counts, installed roles/features, and key migration review items |
| System | Hardware, virtualization, operating system, processor, memory, and boot data |
| Network | IPv4 address, prefix, subnet mask, gateway, DNS, and MAC address by active interface |
| Storage | Fixed-volume capacity and free space |
| Roles and Features | Every installed Windows Server role and feature returned by Server Manager |
| Applications | Installed software from the 32-bit and 64-bit uninstall registry |
| Services | Service state, startup mode, logon identity, and executable path |
| Listening Ports | Listening TCP endpoints with owning process |
| Shares and Permissions | SMB share details, share permissions, and root NTFS permissions |
| Printers | Printer, share, driver, port, and publication details |
| Scheduled Tasks | Non-Microsoft scheduled tasks by default; use `-IncludeMicrosoftTasks` to include all tasks |
| Local Accounts | Local users on member servers; intentionally empty on domain controllers |
| Local Group Membership | Local group membership on member servers; intentionally empty on domain controllers |
| Service Accounts | Non-built-in Windows service identities and their service usage |
| Azure Components | Yes/No detection signals for Azure and Entra agents or applications |
| Active Directory | Local domain controller, domain, forest, site, Global Catalog, RODC, and FSMO data |
| AD Users | Domain user identity, contact, organizational, status, logon, and service-account indicators |
| AD Group Membership | Direct and primary group membership by user |
| AD Replication | Replication partners, last successful replication, and status |
| DHCP Scopes | Scope ranges, exclusions, reservations, options, lease, utilization, NAP, and failover |
| DNS Zones | Forward and reverse zone properties |
| Diagnostics | Collector result, record count, warning/error message, and timestamp |

Sheets that return no records remain present and clearly state `No records returned.` This distinguishes an empty result from an omitted data category.

## Intentionally manual or deferred

These items cannot be discovered reliably from one server without customer-specific systems, credentials, or decisions:

- Migration disposition, business owner, purpose, location, and project notes
- Backup product policy compliance and successful restore evidence
- Warranty and support entitlement
- Application installation media, migration instructions, licensing, contracts, vendor support, and documentation
- Testing, QA, and risk-management decisions
- Recursive file/folder size and inherited ACL analysis
- DNS record export
- Print-share security beyond the data exposed by the local print and SMB collectors
- Environment-wide server enumeration and consolidated multi-server reporting

## Reviewing the result

Before expanding to all servers, verify:

1. The workbook opens without an Excel repair warning.
2. `Discovery Summary` matches the server and reports the expected collector status.
3. `System`, `Network`, and `Storage` match the server.
4. `Roles and Features` matches Server Manager.
5. Applications are sufficiently complete without `Win32_Product`.
6. Share permissions represent only share roots, as intended.
7. AD, DHCP, and DNS detail is complete for the installed roles.
8. Failed or partial collectors are explained in `Diagnostics`.
9. Sensitive output is stored in an approved location and is not committed to a public repository.

## Existing-script decisions

The earlier discovery scripts remain unchanged for comparison. Phase 1 consolidates their useful intent while changing these patterns:

- AD-wide discovery and repeated remote sessions are deferred.
- `Win32_Product` inventory is replaced with uninstall-registry enumeration.
- ImportExcel is bootstrapped from PSGallery at pinned version 7.8.10.
- Normalized row-based worksheets replace wide, duplicated template sections.
- Domain controllers do not label domain accounts as local accounts.
- Collector-level error handling replaces all-or-nothing execution.
- `Export-Excel -Show` is not used, allowing headless execution.

## Next phase

After the revised single-server workbook is reviewed, Phase 2 can add controlled AD computer discovery, include/exclude filters, OU scoping, remoting timeouts, concurrency limits, offline-server reporting, and consolidated multi-server output.
