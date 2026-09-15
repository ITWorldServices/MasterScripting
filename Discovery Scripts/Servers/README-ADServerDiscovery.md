# AD-wide Windows Server discovery

`Invoke-ADServerDiscovery.ps1` inventories enabled Windows Server computer
objects from Active Directory and produces one consolidated workbook in
`C:\Temp`. It does not change or replace the single-server
`Invoke-ServerDiscovery.ps1` workflow.

## Files to download

Keep these three files together in the same directory:

- `Invoke-ADServerDiscovery.ps1` — AD enumeration, remoting, consolidation, and Excel output
- `Get-ServerDiscoveryData.ps1` — data-only collector executed through PowerShell remoting
- `Get-LegacyServerDiscoveryData.ps1` — bounded WMI/DCOM basic-inventory fallback

The collector is sent by `Invoke-Command -FilePath`; it does not need to be
copied to each remote server. Remote servers do not need ImportExcel and do not
connect to GitHub. Only the initiating server creates the workbook.

## Prerequisites

- Windows PowerShell 5.1
- ActiveDirectory PowerShell module on the initiating server
- PowerShell remoting/WinRM allowed from the initiating server to targets
- WMI/DCOM and RPC firewall access to any target that may require fallback collection
- An account with permission to query AD and inventory each target
- HTTPS access to PowerShell Gallery/NuGet on the initiating server if
  ImportExcel 7.8.10 is not already installed

## Run

From an elevated Windows PowerShell session:

```powershell
Set-Location C:\Scripts
.\Invoke-ADServerDiscovery.ps1
```

The default output is:

```text
C:\Temp\<DOMAIN>-ServerDiscovery-<timestamp>.xlsx
```

Useful options:

```powershell
# Limit discovery to a server OU
.\Invoke-ADServerDiscovery.ps1 -SearchBase 'OU=Servers,DC=contoso,DC=com'

# Include or exclude servers with wildcard patterns
.\Invoke-ADServerDiscovery.ps1 -IncludeComputerName 'PROD-*' -ExcludeComputerName '*-OLD','*-DR'

# Supply alternate credentials and adjust parallelism
$credential=Get-Credential
.\Invoke-ADServerDiscovery.ps1 -Credential $credential -ThrottleLimit 12

# Allow up to 90 seconds for the WMI/DCOM fallback phase
.\Invoke-ADServerDiscovery.ps1 -FallbackTimeoutSeconds 90
```

## Collection rules

All reachable servers receive the basic inventory:

- System
- Network
- Storage
- Roles and Features
- Applications
- Services
- Listening Ports
- Shares and Permissions
- Scheduled Tasks
- Local Accounts
- Local Group Membership
- Service Accounts
- Diagnostics

SQL Server and Exchange Server are detected from installed Windows services and
are classified as **Basic only**. They remain in the consolidated workbook, but
AD DS, DHCP, and DNS collectors are skipped so separate SQL and Exchange
discovery scripts can be added later.

General servers run role-specific collectors only when the role is installed:

| Installed role | Additional discovery |
| --- | --- |
| AD Domain Services | Domain controller, AD users, group membership, replication |
| DHCP Server | Scopes, exclusions, reservations, options, failover |
| DNS Server | DNS zones |

AD users and group memberships are de-duplicated in the consolidated workbook
when multiple domain controllers return the same domain data.

## WMI/DCOM fallback

If any target does not return a modern PowerShell payload, the orchestrator
retries it over WMI/DCOM from the initiating server. This covers legacy Windows
Server versions as well as newer systems where WinRM is disabled, incompatible,
or temporarily unavailable. A successful fallback restores the server's basic
rows instead of leaving it absent from the detail sheets.

The fallback collects system, network, storage, server features, applications,
services, shares, scheduled tasks, local accounts, local group membership, and
service accounts where the operating system exposes them. Legacy operating
systems are identified as `Legacy basic only`; newer systems are identified as
`WMI fallback basic only` on `Server Summary`.

Listening ports and expanded share/NTFS permissions are not reliably available
through WMI/DCOM and are recorded as skipped or limited in `Diagnostics`.
Role-specific AD, DHCP, and DNS discovery remains on the modern collector.

Fallback attempts run as background jobs in parallel. The default 60-second
timeout applies to the complete fallback phase, not to each individual WMI
query. A target that exceeds the timeout is stopped, marked unavailable, and
does not prevent the workbook from being created. Use
`-FallbackTimeoutSeconds` to adjust the limit between 15 and 900 seconds.
`-LegacyTimeoutSeconds` remains accepted as a compatibility alias.

## Failures and unavailable servers

An offline, inaccessible, or remoting-disabled server does not stop the run.
Every AD target appears on `Server Summary`; its status and error are recorded
there and on `Diagnostics`. Collector-level failures are also retained, while
successful data from the rest of the environment is still exported.

Primary WinRM and fallback errors are also printed in the console. Primary
errors retain the error ID and available source location in `Diagnostics`,
including when WMI fallback subsequently succeeds.

## Updating the collectors

Download all three `.ps1` files listed above together when applying a fix.
The diagnostics-array fix changes both `Get-ServerDiscoveryData.ps1` and
`Get-LegacyServerDiscoveryData.ps1`; replacing only the orchestrator does not
apply it. Both collectors now convert their diagnostics lists with `ToArray()`
to prevent an array-conversion error from discarding an otherwise collected
inventory. The remote collector still supports PowerShell 4.0.

The payload regression check uses synthetic data and makes no AD, WinRM, WMI,
or Excel calls. From the repository's server-discovery directory, run:

```powershell
.\Tests\Test-DiscoveryPayload.ps1
```
