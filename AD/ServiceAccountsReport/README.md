# Service Accounts Report Scripts

This directory contains two PowerShell scripts for auditing service accounts across Windows servers:

1. **ServiceAccountsReport.ps1** - Basic service account collection
2. **ServiceAccountsReportAndAnalysis.ps1** - Enhanced security analysis and reporting

## Overview

These scripts are designed to remotely collect information about services and the accounts they run under on Windows servers. This is particularly useful for security auditing, compliance checking, and identifying potential security risks.

## Scripts Description

### ServiceAccountsReport.ps1 (Basic Version)

A simple script that collects basic service information from remote servers.

**Features:**
- Connects to servers from a text file list
- Queries all services on each server
- Collects service name, display name, start mode, state, and service account
- Exports results to CSV format

**Use Case:** Quick service account inventory without advanced analysis.

### ServiceAccountsReportAndAnalysis.ps1 (Enhanced Version)

A comprehensive security analysis tool with advanced features and reporting.

**Features:**
- **Parallel Processing**: Scan multiple servers simultaneously for faster execution
- **Active Directory Integration**: Lookup account details, password age, group memberships
- **Security Analysis**: Identify high-risk configurations and suspicious patterns
- **CIS Compliance**: Check against CIS benchmark controls
- **Advanced Filtering**: Focus on administrative accounts or include all services
- **Multiple Output Formats**: CSV and HTML reports with executive summary
- **Error Handling**: Comprehensive logging and connectivity testing
- **Process Information**: Memory usage, binary paths, company information

## Prerequisites

### Basic Requirements (Both Scripts)
- PowerShell 5.1 or later
- Administrative permissions on target servers
- Network connectivity to remote servers
- WMI access enabled on target servers (port 135 and dynamic RPC ports)

### Additional Requirements (Enhanced Script)
- **Active Directory Module** (optional, for AD lookups)
  ```powershell
  Install-WindowsFeature RSAT-AD-PowerShell
  # or
  Install-Module ActiveDirectory
  ```
- **Sufficient privileges** for AD queries (if using -EnableADLookup)

## Setup

1. **Create Server List File**
   
   Create a `servers.txt` file in the same directory as the scripts:
   ```
   SERVER01
   SERVER02.domain.com
   10.1.1.100
   ```

2. **Verify Connectivity**
   
   Test WMI connectivity to your servers:
   ```powershell
   Get-WmiObject -Class Win32_OperatingSystem -ComputerName "SERVER01"
   ```

## Usage

### Basic Script (ServiceAccountsReport.ps1)

**Simple Execution:**
```powershell
.\ServiceAccountsReport.ps1
```

This will:
- Read servers from `servers.txt`
- Scan all services on each server
- Export results to `ServiceAccountsReport.csv`

### Enhanced Script (ServiceAccountsReportAndAnalysis.ps1)

**Basic Usage:**
```powershell
.\ServiceAccountsReportAndAnalysis.ps1
```

**Advanced Usage Examples:**

```powershell
# Comprehensive scan with all features enabled
.\ServiceAccountsReportAndAnalysis.ps1 -ServerListPath "prod_servers.txt" -OutputPath ".\SecurityReports" -EnableADLookup -IncludeNonAdminServices -MaxParallelJobs 10

# Quick security-focused scan (admin accounts only)
.\ServiceAccountsReportAndAnalysis.ps1 -MaxParallelJobs 5

# Scan with custom timeout and include all services
.\ServiceAccountsReportAndAnalysis.ps1 -TimeoutSeconds 60 -IncludeNonAdminServices

# Scan specific server list with AD integration
.\ServiceAccountsReportAndAnalysis.ps1 -ServerListPath "critical_servers.txt" -EnableADLookup -OutputPath "C:\Reports"
```

## Parameters (Enhanced Script)

| Parameter | Type | Default | Description |
|-----------|------|---------|-------------|
| `ServerListPath` | String | "servers.txt" | Path to file containing server names (one per line) |
| `OutputPath` | String | ".\Reports" | Directory for output files |
| `IncludeNonAdminServices` | Switch | False | Include services not running under administrative accounts |
| `TimeoutSeconds` | Integer | 30 | Connection timeout for each server |
| `MaxParallelJobs` | Integer | 5 | Maximum concurrent server scans |
| `EnableADLookup` | Switch | False | Enable Active Directory account lookups |

## Output Files

### Basic Script Output
- **ServiceAccountsReport.csv**: Simple CSV with basic service information.

### Enhanced Script Output
- **ServiceAccountsReport_[timestamp].csv**: Detailed CSV with all collected data
- **ServiceAccountsReport_[timestamp].html**: Interactive HTML report with:
  - Executive summary with key metrics
  - High-risk services highlighted
  - Sortable tables with all service details
  - Color-coded risk levels
- **ServiceScanErrors.log**: Log of any connection or scanning errors

## Security Analysis Features

The enhanced script performs several security checks:

### High-Risk Account Detection
- Domain admin accounts running services
- Enterprise admin accounts
- Schema admin accounts
- Suspicious service account patterns

### Suspicious Configuration Detection
- Services running from user directories
- Unusual executable extensions (.tmp, .bat, .vbs, .ps1)
- Services in suspicious paths (temp, downloads, etc.)
- Services running while disabled

### CIS Compliance Checks
- **CIS Control 5.1**: Unnecessary services detection (Telnet, SNMP, Spooler, Fax, RemoteRegistry)
- **CIS Control 9.1**: Services running with excessive privileges
- Manual services currently running

### Active Directory Integration
When `-EnableADLookup` is enabled, the script provides:
- Account enabled/disabled status
- Last logon information
- Password age in days
- Group memberships
- Service account identification

## Report Interpretation

### Risk Levels in HTML Report
- **High Risk** (Red): Services with security flags - immediate attention required
- **Medium Risk** (Orange): CIS non-compliant services - review recommended
- **Low Risk** (Green): Compliant services

### Key Metrics to Monitor
- Services running under domain admin accounts
- Services with disabled accounts
- Services in user directories
- Unnecessary services that are running
- Accounts with old passwords (if AD lookup enabled)

## Troubleshooting

### Common Issues

**WMI Connection Failures:**
- Verify Windows Firewall settings
- Check if Remote Registry service is running
- Ensure DCOM permissions are correct

**AD Lookup Failures:**
- Verify Active Directory module is installed
- Check domain connectivity
- Ensure account has read permissions to AD

**Performance Issues:**
- Reduce `MaxParallelJobs` for slower networks
- Increase `TimeoutSeconds` for slow servers
- Use basic script for simple inventories

### Error Logs
Check `ServiceScanErrors.log` in the output directory for detailed error information.

## Best Practices

1. **Start Small**: Test with a few servers before running against entire infrastructure
2. **Use Filtering**: Use basic script or disable `-IncludeNonAdminServices` for focused security reviews
3. **Regular Scanning**: Schedule weekly or monthly scans for security monitoring
4. **Review HTML Reports**: Use the interactive HTML report for analysis and presentations
5. **Monitor Changes**: Compare reports over time to identify configuration drift

## Security Considerations

- Scripts require administrative access to target servers
- Consider running from a secure administrative workstation
- Protect output files as they contain sensitive security information
- Review and validate findings before taking corrective actions

## Examples of Security Findings

The enhanced script can identify issues such as:
- Critical services running under domain admin accounts
- Services executing from temporary directories
- Disabled accounts still being used by running services
- Non-standard services with suspicious executable paths
- Services with excessive privileges per CIS benchmarks

## Support

For issues or questions:
1. Check the error log file for specific error messages
2. Verify prerequisites are met
3. Test connectivity manually using WMI commands
4. Review PowerShell execution policy settings
