# Default values
$siteUrlDefault = "https://blackhawkcenter.sharepoint.com/teams/GoodwinAustinTeam"
$groupEmailDefault = "GoodwinAustinTeamAccess@goodwintx.com"
$clientIdDefault = "5e940a58-2d36-4e9e-9d6f-47009e181aea"
$permissionLevelDefault = "Contribute"

# Permission level options
$permissionLevels = @("Read", "Contribute", "Edit", "Full Control")

# Prompt for URL
$siteUrl = Read-Host "Enter the SharePoint Site URL [$siteUrlDefault]"
if ([string]::IsNullOrWhiteSpace($siteUrl)) { $siteUrl = $siteUrlDefault }

# Prompt for group email
$groupEmail = Read-Host "Enter the Group Email Address [$groupEmailDefault]"
if ([string]::IsNullOrWhiteSpace($groupEmail)) { $groupEmail = $groupEmailDefault }

# Prompt for permission level
Write-Host "Select Permission Level:"
for ($i = 0; $i -lt $permissionLevels.Count; $i++) {
    Write-Host "$($i+1): $($permissionLevels[$i])"
}
$permChoice = Read-Host "Enter the number for permission level [$permissionLevelDefault]"
if ([string]::IsNullOrWhiteSpace($permChoice)) {
    $permissionLevel = $permissionLevelDefault
} else {
    $permIndex = [int]$permChoice - 1
    if ($permIndex -ge 0 -and $permIndex -lt $permissionLevels.Count) {
        $permissionLevel = $permissionLevels[$permIndex]
    } else {
        Write-Host "Invalid choice; using default permission level."
        $permissionLevel = $permissionLevelDefault
    }
}

# Prompt for Client ID
$clientId = Read-Host "Enter Client ID [$clientIdDefault]"
if ([string]::IsNullOrWhiteSpace($clientId)) { $clientId = $clientIdDefault }

# Connect
Connect-PnPOnline -Url $siteUrl -ClientId $clientId -Interactive

# Ensure the group is registered in the site user info list
$principal = Get-PnPUser -Identity $groupEmail -ErrorAction SilentlyContinue

if ($null -eq $principal) {
    # Add to site members group (will register the principal)
    $membersGroup = Get-PnPGroup | Where-Object { $_.Title -match "Members" }
    if ($membersGroup) {
        Add-PnPGroupMember -LoginName $groupEmail -Identity $membersGroup.Id
        Write-Host "Added $groupEmail to Site Members group to register it."
        # Try to resolve again
        $principal = Get-PnPUser -Identity $groupEmail -ErrorAction SilentlyContinue
    }
}

# Get all document libraries except Style Library
$libraries = Get-PnPList | Where-Object { $_.BaseTemplate -eq 101 -and $_.Title -ne "Style Library" }

# Display libraries to be modified
Write-Host "`nThe following document libraries will be processed (excluding 'Style Library'):"
$libraries | ForEach-Object { Write-Host " - $($_.Title)" }

# Prompt to continue
Read-Host "`nReview the list above. Press Enter to continue or Ctrl+C to cancel"

foreach ($lib in $libraries) {
    # Break inheritance for unique permissions on the library
    Set-PnPList -Identity $lib -BreakRoleInheritance

    # Grant group permission to the library
    Set-PnPListPermission -Identity $lib -User $groupEmail -AddRole $permissionLevel

    # Ensure group exists in SharePoint's user info list
    $principal = Get-PnPUser -Identity $groupEmail
    if ($null -ne $principal) {
        # Get permissions for that principal id
        $perm = Get-PnPListPermissions -Identity $lib -PrincipalId $principal.Id
        if ($perm.RoleDefinitions -contains $permissionLevel) {
            Write-Host "SUCCESS: '$permissionLevel' assigned to $groupEmail on '$($lib.Title)'."
        } else {
            Write-Host "FAILED: '$permissionLevel' NOT assigned to $groupEmail on '$($lib.Title)'."
        }
    } else {
        Write-Host "FAILED: Could not resolve principal for $groupEmail on '$($lib.Title)'."
    }
}
