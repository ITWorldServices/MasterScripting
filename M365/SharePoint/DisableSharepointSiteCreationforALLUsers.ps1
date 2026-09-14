# Connect to SharePoint Online
Connect-SPOService -Url https://icsholdingllc-admin.sharepoint.com

# Get the current organization settings
$orgSettings = Get-SPOTenant

# Disable self-service site creation
$orgSettings.DisableSharingForNonOwners = $true
$orgSettings.Update()

# Disable self-service site creation at the user level
$users = Get-SPOSiteUser -Site https://icsholdingllc.sharepoint.com
foreach ($user in $users) {
    Set-SPOSite -Identity $user.SiteUrl -SharingCapability Disabled
}

Write-Host "Self-service site creation has been disabled for all users."

# Disconnect from SharePoint Online
#Disconnect-SPOService