# Standalone local server discovery

Download **only `Invoke-StandaloneServerDiscovery.ps1`** and copy it onto the
server you want to inventory. This is the individual-server companion to the
AD-wide discovery workflow. It works for domain members and workgroup servers,
including a server that cannot be inventoried remotely from the AD server.

## Run on the individual server

1. Sign in to that server with its appropriate administrator account. This can
   be a local administrator on a workgroup server or a domain administrator.
2. Copy `Invoke-StandaloneServerDiscovery.ps1` to `C:\Scripts`.
3. Open **Windows PowerShell as administrator**. Use 64-bit PowerShell on a
   64-bit server so the application inventory can read both registry views.
4. Run:

```powershell
Set-Location C:\Scripts
.\Invoke-StandaloneServerDiscovery.ps1
```

If an Internet-download mark blocks the file, use its Windows file properties
to unblock it according to your environment's execution policy.

The script uses the account running the elevated session. It does not accept,
store, or prompt for passwords, or change remote-access permissions. Running it
on your AD server inventories the AD server; to inventory TABS3SRV, run this
file **on TABS3SRV itself**.

## Output

Both files are written locally:

```text
C:\Temp\SERVERNAME-ServerDiscovery-yyyyMMdd-HHmmss-fff.xlsx
C:\Temp\SERVERNAME-ServerDiscovery-yyyyMMdd-HHmmss-fff.clixml
```

The XLSX is the reviewable report. The CLIXML is the complete raw inventory
snapshot and is saved first, retaining data if Excel export fails or a value
exceeds Excel's cell-length limit. Neither file is uploaded anywhere.

Use another destination or include Microsoft scheduled tasks:

```powershell
.\Invoke-StandaloneServerDiscovery.ps1 -OutputDirectory D:\Discovery
.\Invoke-StandaloneServerDiscovery.ps1 -IncludeMicrosoftTasks
```

## Collection and requirements

- Windows Server 2008 R2 or later, with Windows PowerShell 2.0 or later.
- An elevated administrator session on the server being inventoried.
- No domain membership, WinRM configuration, remote WMI access, GitHub
  connection, Excel installation, ImportExcel, NuGet, or Gallery access required.
- XLSX packaging uses .NET classes supplied with PowerShell. The embedded C#
  helper is compiled in memory by `Add-Type`; no downloaded executable is used.
- Inventory queries are read-only. Output files are the only deliberate writes.
- No AD computer enumeration, application repair (`Win32_Product`), recursive
  file-share size scans, or changes to services, firewall, or authentication.
- AD/DHCP/DNS queries on eligible role hosts can still require connectivity to
  their normal directory/services; no network-isolated AD behavior is promised.

On Windows Server 2012 or later with PowerShell 4.0 or later, the script runs
an embedded copy of the current `Get-ServerDiscoveryData.ps1` collector locally.
It follows the same SQL/Exchange classification and installed-role rules as the
AD-wide workflow. SQL/Exchange hosts receive basic inventory only. AD, DHCP,
and DNS discovery runs only when the corresponding role is detected.

On Windows Server 2008 R2, or with PowerShell 2.0/3.0, it automatically uses
**Local legacy basic only** collection. This uses local WMI, registry reads,
Task Scheduler COM, netstat, and WinNT/local ACL interfaces. It avoids the
remote fallback's per-property registry calls and 60-second phase timeout.
It reports each section as it starts. A local OS/provider call can still be
slow; this version does not enforce a per-section hard timeout.

For troubleshooting on newer servers, force that local basic path with:

```powershell
.\Invoke-StandaloneServerDiscovery.ps1 -UseLegacyCollector
```

## Workbook coverage

The workbook uses the AD-wide field names and section names. It includes
Server Summary, System, Network, Storage, Roles and Features, Applications,
Services, Listening Ports, Shares and Permissions, Scheduled Tasks, Local
Accounts, Local Group Membership, Service Accounts, the AD/DHCP/DNS sections,
and Diagnostics. It does not append to an existing AD-wide workbook.

Empty, failed, and unsupported sections remain visible. Diagnostics records
the reason; blank sections do not imply that the server has no such data.

Legacy limitations:

- AD/DHCP/DNS detail sections are retained but explicitly skipped. General
  legacy hosts receive basic inventory even if they have those roles installed.
- Roles/features contain what `Win32_ServerFeature` exposes.
- Local-account LastLogon is unavailable through `Win32_UserAccount` and is
  explicitly reported as blank. Local accounts/groups are skipped on DCs.
- Share and share-root NTFS permissions are attempted locally. Unreadable ACLs
  generate warnings; subdirectories are not traversed.
- Non-executable scheduled-task actions are recorded by action type.

The current service-based SQL classifier also recognizes Windows Internal
Database service names. This is the same known classification limitation as
the AD-wide collector; WID-only hosts are still included in basic inventory.

## Maintenance and validation

The downloadable script is generated and contains all its dependencies.
Maintainers edit `Standalone/*.ps1`, `Standalone/ZipWriter.cs`, and the shared
modern collector, then rebuild from a PowerShell 5.1+ workstation:

```powershell
.\Standalone\Build-StandaloneServerDiscovery.ps1
.\Tests\Test-StandaloneDiscovery.ps1
.\Tests\Test-DiscoveryPayload.ps1
```

The build check prevents the embedded modern collector from drifting from its
source. Tests cover runtime selection, legacy-compatible syntax restrictions,
local WMI arguments, partial failures, raw snapshots, workbook structure and
typed values. Windows-specific WMI/COM/AD calls and execution under Windows
PowerShell 2.0 still require validation on representative servers. Portable
tests do not establish that each target grants every inventory query.
