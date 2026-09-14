<#
.SYNOPSIS
Retrieves and exports all DNS service configuration records for every domain in a Microsoft 365 tenant (Commercial or GCC High) using Microsoft Graph PowerShell, outputting results to a CSV file with color-coded console display for easy review.

.REQUIREMENTS
- PowerShell 5.1 or later
- Microsoft Graph PowerShell module (Microsoft.Graph)
- Microsoft 365 tenant with Commercial or GCC High environment access
- Sufficient permissions to read domain and DNS configuration records (Domain.Read.All scope)
- Network access to Microsoft 365 services
- Write access to C:\temp directory

.SUMMARY
This script automates the process of connecting to Microsoft Graph in a Microsoft 365 tenant (Commercial or GCC High), retrieving all domains, and exporting their DNS service configuration records. It installs the Microsoft Graph module if missing, connects with the required permissions, and fetches all domain-related DNS records. The results are displayed in the console with color-coded formatting for enhanced readability and are also exported to a timestamped CSV file in the C:\temp directory. This tool is especially useful for administrators managing DNS records, troubleshooting, or preparing for tenant migrations in secure or commercial cloud environments.

.SYNTAX
.\Get-TenantDnsRecords.ps1

# No parameters required. 
# The script will prompt for authentication if needed.
# Output CSV will be saved to C:\temp\[TenantName]_O365TenantRecords_[Timestamp].csv

#>

CLEAR

# Prompt user to select tenant environment
Write-Host "Select Microsoft 365 tenant environment:" -ForegroundColor Cyan
Write-Host "1. Commercial" -ForegroundColor yellow
Write-Host "2. GCC High" -ForegroundColor red
$tenantChoice = Read-Host "Enter 1 for Commercial or 2 for GCC High"

switch ($tenantChoice) {
    '1' {
        $mgEnvironment = "Global"
        Write-Host "You selected Commercial tenant." -ForegroundColor Green
    }
    '2' {
        $mgEnvironment = "USGov"
        Write-Host "You selected GCC High tenant." -ForegroundColor Green
    }
    Default {
        Write-Host "Invalid selection. Exiting script." -ForegroundColor Red
        exit
    }
}

# Install Microsoft Graph module if not already installed
if (-not (Get-Module -ListAvailable -Name Microsoft.Graph)) {
    Install-Module Microsoft.Graph -Scope CurrentUser -Force
}

# Connect to Microsoft Graph for selected environment
Connect-MgGraph -Environment $mgEnvironment -Scopes "Domain.Read.All"

# Get all domains in the tenant
$domains = Get-MgDomain

# Prepare output list
$results = @()

# Get tenant name and sanitize for filename
$tenantName = (Get-MgOrganization).DisplayName
$tenantNameSanitized = $tenantName -replace '[^a-zA-Z0-9_-]', '_'

# Get timestamp
$timestamp = Get-Date -Format "yyyyMMddHHmmss"

foreach ($domain in $domains) {
    try {
        $records = Get-MgDomainServiceConfigurationRecord -DomainId $domain.Id
        if ($records) {
            foreach ($record in $records) {
                # Prepare properties for CSV
                $propString = ($record.AdditionalProperties | Out-String).Trim()
                # Add to results for CSV
                $obj = [PSCustomObject]@{
                    TenantName = $tenantName
                    Timestamp = $timestamp
                    Domain = $domain.Id
                    RecordType = $record.RecordType
                    Label = $record.Label
                    SupportedService = $record.SupportedService
                    Properties = $propString
                }
                $results += $obj
                # Display in console with color
                Write-Host "`nDomain: " -NoNewline
                Write-Host $domain.Id -ForegroundColor Red
                Write-Host "Type: $($record.RecordType)"
                Write-Host "Label: $($record.Label)"
                Write-Host "Supported Service: $($record.SupportedService)"
                Write-Host "Properties:"
                foreach ($kv in $record.AdditionalProperties.GetEnumerator()) {
                    Write-Host ("  {0}: " -f $kv.Key) -NoNewline -ForegroundColor Yellow
                    Write-Host $kv.Value -ForegroundColor Green
                }
            }
        }
    } catch {
        Write-Host "Error retrieving records for $($domain.Id): $_" -ForegroundColor Red
    }
}

# Create output directory if missing
if (-not (Test-Path "C:\temp")) {
    New-Item -ItemType Directory -Path "C:\temp" | Out-Null
}

# Export to CSV with tenant name prepended
$outputPath = "C:\temp\${tenantNameSanitized}_O365TenantRecords_${timestamp}.csv"
$results | Export-Csv -Path $outputPath -NoTypeInformation -Encoding UTF8

Write-Host "`nSuccessfully exported $($results.Count) DNS records to:"
Write-Host $outputPath -ForegroundColor Green

# Disconnect when done
Disconnect-MgGraph
