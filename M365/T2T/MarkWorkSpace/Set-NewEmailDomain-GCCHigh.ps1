# Connect to Azure AD in GCC High
Connect-AzureAD -AzureEnvironmentName AzureUSGovernment

# Get all users
$allUsers = Get-AzureADUser -All $true

# Get all admin roles
$adminRoles = Get-AzureADDirectoryRole

# Collect all admin user ObjectIds
$adminUserIds = @()
foreach ($role in $adminRoles) {
    $members = Get-AzureADDirectoryRoleMember -ObjectId $role.ObjectId
    $adminUserIds += $members.ObjectId
}
$adminUserIds = $adminUserIds | Select-Object -Unique

# Filter non-admin users
$nonAdminUsers = $allUsers | Where-Object { $adminUserIds -notcontains $_.ObjectId }

# Define your new UPN suffix
$newSuffix = "i3dmfg.com"  # <-- Replace with your verified domain

# Change UPN for each non-admin user
foreach ($user in $nonAdminUsers) {
    $oldUPN = $user.UserPrincipalName
    $username = $oldUPN.Split("@")[0]
    $newUPN = "$username@$newSuffix"

    # Skip if already correct
    if ($oldUPN -ne $newUPN) {
        Write-Host "Changing UPN for $oldUPN to $newUPN"
        Set-AzureADUser -ObjectId $user.ObjectId -UserPrincipalName $newUPN
    }
}
