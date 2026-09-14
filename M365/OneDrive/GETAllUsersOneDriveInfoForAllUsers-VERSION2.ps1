<#
.SYNOPSIS
This script generates a OneDrive usage report for all users in a Microsoft 365 tenant.

.DESCRIPTION
The script performs the following actions:
1. Connects to Microsoft Graph API with necessary permissions.
2. Retrieves all users in the tenant.
3. For each user, it attempts to get their OneDrive usage details.
4. If OneDrive is provisioned, it records the usage data.
5. If OneDrive is not provisioned, it marks the user accordingly.
6. Exports the results to a CSV file.
7. Displays the results in the console.
8. Provides a summary of total users, provisioned, and non-provisioned OneDrives.

.NOTES
- Requires the Microsoft Graph PowerShell SDK.
- Requires "User.Read.All" and "Sites.Read.All" permissions.
- CSV file is saved at "C:\TEMP\AR-OneDriveUsageReport.csv".
- Displays progress for each user in the console.

.EXAMPLE
.\Get-OneDriveUsageReport.ps1

#>

Import-Module Microsoft.Graph.Authentication
Import-Module Microsoft.Graph.Sites
Import-Module Microsoft.Graph.Users

# Connect to Microsoft Graph
Connect-MgGraph -Scopes "User.Read.All", "Sites.Read.All"

# Get all users
$users = Get-MgUser -All

# Initialize an array to store the results
$results = @()

foreach ($user in $users) {
    try {
        # Get OneDrive site for the user
        $oneDriveSite = Get-MgUserDrive -UserId $user.Id -ErrorAction Stop

        # Get OneDrive usage details
        $driveUsage = Get-MgDrive -DriveId $oneDriveSite.Id

        $result = [PSCustomObject]@{
            User = $user.UserPrincipalName
            TotalSizeMB = [math]::Round($driveUsage.Quota.Used / 1MB, 2)
            ProvisioningStatus = "Provisioned"
            LastAccessTime = $driveUsage.LastModifiedDateTime
        }
    }
    catch {
        # If OneDrive is not provisioned, add user with empty data
        $result = [PSCustomObject]@{
            User = $user.UserPrincipalName
            TotalSizeMB = 0
            ProvisioningStatus = "Not Provisioned"
            LastAccessTime = $null
        }
    }

    $results += $result
    Write-Host "$($user.UserPrincipalName) - Used Storage: $($result.TotalSizeMB) MB"
}

# Export the results to a CSV file
$results | Export-Csv -Path "C:\TEMP\AR-OneDriveUsageReport.csv" -NoTypeInformation

# Display the results in the console
$results | Format-Table -AutoSize

# Display summary
$totalUsers = $results.Count
$provisionedUsers = ($results | Where-Object { $_.ProvisioningStatus -eq "Provisioned" }).Count
$notProvisionedUsers = $totalUsers - $provisionedUsers

Write-Host "`nSummary:"
Write-Host "Total Users: $totalUsers"
Write-Host "Provisioned OneDrives: $provisionedUsers"
Write-Host "Not Provisioned OneDrives: $notProvisionedUsers"

Disconnect-MgGraph