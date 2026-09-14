<#
.SYNOPSIS
Exports SharePoint Online site members to a CSV file with tenant name and timestamp in the filename.

.DESCRIPTION
This script connects to SharePoint Online using the provided or previously used admin URL,
retrieves all site collections, and enumerates their members. The results are exported to a CSV file 
in the C:\temp directory. The export file is named using the tenant name (extracted from the admin URL)
and a timestamp in the format YYYYMMDD_HHMMSS to ensure uniqueness and traceability.

The script also remembers the last used admin URL for convenience and prompts the user to reuse it or enter a new one.

.NOTES
Author: Mark Lafnitzegger
Date: July 2025
Requires: SharePoint Online Management Shell
Output: CSV file saved to C:\temp\

.EXAMPLE
.\Export-SPOMembers.ps1

Prompts for the SharePoint Online admin URL, connects to the service, retrieves site members, and exports them to:
C:\temp\contoso_SharePointSiteMembers_20250714_162045.csv
#>

# File to store the last used SharePoint admin URL
$LastUrlFile = "C:\temp\LastSPOAdminUrl.txt"

# Function to get the SharePoint admin URL from user
function Get-AdminUrl {
    if (Test-Path $LastUrlFile) {
        $LastUrl = Get-Content $LastUrlFile
        $Prompt = "Press Enter to use the last SharePoint admin URL (`"$LastUrl`") or enter a new one:"
        $AdminUrl = Read-Host $Prompt
        if ([string]::IsNullOrWhiteSpace($AdminUrl)) {
            $AdminUrl = $LastUrl
        }
    } else {
        $AdminUrl = Read-Host "Enter your SharePoint Online admin URL (e.g., https://contoso-admin.sharepoint.com):"
    }
    # Save the URL for next time
    $AdminUrl | Set-Content $LastUrlFile
    return $AdminUrl
}

# Get the admin URL
$AdminUrl = Get-AdminUrl

# Connect to SharePoint Online
Connect-SPOService -Url $AdminUrl

# Extract tenant name from the URL (e.g., "contoso" from "https://contoso-admin.sharepoint.com")
if ($AdminUrl -match "https:\/\/([a-zA-Z0-9\-]+)-admin\.sharepoint\.com") {
    $TenantName = $matches[1]
} else {
    $TenantName = "UnknownTenant"
}

# Get current timestamp in format YYYYMMDD_HHMMSS
$TimeStamp = Get-Date -Format "yyyyMMdd_HHmmss"

# Construct the export file path
$ExportPath = "C:\temp\${TenantName}_SharePointSiteMembers_${TimeStamp}.csv"

# Get all SharePoint sites
$AllSites = Get-SPOSite -Limit All

# Prepare result array
$Result = @()

# Iterate through each site and get members
foreach ($Site in $AllSites) {
    try {
        $Users = Get-SPOUser -Site $Site.Url
        foreach ($User in $Users) {
            $Result += [PSCustomObject]@{
                SiteUrl = $Site.Url
                Member  = $User.LoginName
            }
        }
    }
    catch {
        Write-Host "Error occurred for $($Site.Url)" -ForegroundColor Yellow
        Write-Host $_ -ForegroundColor Red
    }
}

# Export to CSV with dynamic file name
$Result | Select-Object SiteUrl, Member | Export-Csv -Path $ExportPath -NoTypeInformation -Encoding UTF8

Write-Host "Export complete. File saved to $ExportPath" -ForegroundColor Green
