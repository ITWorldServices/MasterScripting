# EXPORTS ALL PROVISIONED ONEDRIVE ACCOUNTS WITH TOTAL SIZE AND LAST TIME ACCESSED. ALSO REPORTS ON NON-PROVISIONED ACOUNTS.

# Install and import the Microsoft Graph PowerShell SDK if not already installed
# Install-Module Microsoft.Graph -Scope CurrentUser
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
            TotalSizeGB = 0
            ItemCount = 0
            ProvisioningStatus = "Not Provisioned"
            LastAccessTime = $null
        }
    }

    $results += $result
}

# Export the results to a CSV file
$results | Export-Csv -Path "C:\TEMP\GEM-OneDriveUsageReport.csv" -NoTypeInformation

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
