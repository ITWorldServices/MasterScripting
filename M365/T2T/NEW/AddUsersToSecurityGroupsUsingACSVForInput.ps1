<#
.SYNOPSIS
This script adds users to specified security groups in Azure AD using Microsoft Graph.

.DESCRIPTION
This PowerShell script performs the following tasks:
1. Reads a CSV file containing user principal names and corresponding security group names.
     a. Please ensure a header named userPrincipalName and a header named SecurityGroup
2. For each user, it checks if the specified security group exists.
3. If the group exists, it verifies if the user exists.
4. If both user and group exist, it checks if the user is already a member of the group.
5. If the user is not a member, it adds the user to the group.
6. All actions and results are logged to both the console and a CSV output file.
7. The output file is named with the tenant name and a timestamp for easy identification.

.NOTES
- Requires the Microsoft.Graph.Users and Microsoft.Graph.Groups modules.
- Requires appropriate permissions to read and modify users and groups in Azure AD.
- Input CSV should have columns: userPrincipalName, SecurityGroup
#>

Clear

$MaximumFunctionCount = 32768
Import-Module Microsoft.Graph.Users
Import-Module Microsoft.Graph.Groups

# Connect to Microsoft Graph
Connect-MgGraph -Scopes "Group.ReadWrite.All", "User.ReadWrite.All"

# Read the CSV file containing user and group information
$users = Import-Csv -Path "C:\temp\users.csv"

# Create an array to store the output data
$outputData = @()

# Main processing loop
foreach ($user in $users) {
    $userPrincipalName = $user.userPrincipalName
    $groupName = $user.SecurityGroup

    # Section 1: Get the group ID
    $group = Get-MgGroup -Filter "displayName eq '$groupName'"
    if ($null -eq $group) {
        $message = "Group '$groupName' not found for user '$userPrincipalName'. Skipping..."
        Write-Host -ForegroundColor Red $message
        $outputData += [PSCustomObject]@{
            UserPrincipalName = $userPrincipalName
            SecurityGroup = $groupName
            Status = "Error"
            Message = $message
        }
        continue
    }
    $groupId = $group.Id
    $message = "The ID of the group '$groupName' is: $groupId"
    Write-Host $message

    # Section 2: Get the user ID
    $userObj = Get-MgUser -Filter "userPrincipalName eq '$userPrincipalName'"
    if ($null -eq $userObj) {
        $message = "User '$userPrincipalName' not found. Skipping..."
        Write-Host -ForegroundColor Red $message
        $outputData += [PSCustomObject]@{
            UserPrincipalName = $userPrincipalName
            SecurityGroup = $groupName
            Status = "Error"
            Message = $message
        }
        continue
    }
    $userId = $userObj.Id
    $message = "The ID of the user '$userPrincipalName' is: $userId"
    Write-Host $message

    # Section 3: Check if the user is already a member of the group
    $existingMember = Get-MgGroupMember -GroupId $groupId | Where-Object { $_.Id -eq $userId }
    if ($null -ne $existingMember) {
        $message = "Already a member of $groupName."
        Write-Host -ForegroundColor Yellow $message
        $outputData += [PSCustomObject]@{
            UserPrincipalName = $userPrincipalName
            SecurityGroup = $groupName
            Status = "Already Member"
            Message = $message
        }
        continue
    }

    # Section 4: Add the user to the group
    try {
        New-MgGroupMember -GroupId $groupId -DirectoryObjectId $userId
        $message = "Successfully added user '$userPrincipalName' to group '$groupName'"
        Write-Host -ForegroundColor Green $message
        $outputData += [PSCustomObject]@{
            UserPrincipalName = $userPrincipalName
            SecurityGroup = $groupName
            Status = "Success"
            Message = $message
        }
    }
    catch {
        $message = "Failed to add user '$userPrincipalName' to group '$groupName': $_"
        Write-Host -ForegroundColor Red $message
        $outputData += [PSCustomObject]@{
            UserPrincipalName = $userPrincipalName
            SecurityGroup = $groupName
            Status = "Error"
            Message = $message
        }
    }
}

# Section 5: Generate output file name with tenant name and timestamp
# Get the tenant information
$tenantInfo = Get-MgOrganization
$tenantName = $tenantInfo.DisplayName

# Clean up the tenant name to make it suitable for a file name
$cleanTenantName = $tenantName -replace '[^\w\-\.]', '_'

# Generate a timestamp for the file name
$timestamp = Get-Date -Format "yyyyMMdd_HHmmss"

# Export the output data to a CSV file with timestamp and tenant name in the filename
$outputFilePath = "C:\temp\output_{0}_{1}.csv" -f $cleanTenantName, $timestamp
$outputData | Export-Csv -Path $outputFilePath -NoTypeInformation

Write-Host "Output has been saved to: $outputFilePath"

# Disconnect from Microsoft Graph
Disconnect-MgGraph