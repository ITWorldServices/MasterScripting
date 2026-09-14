# ADD SITE ADMIN to ALL SHAREPOINT SITES
# Load SharePoint Online Management Shell
Import-Module Microsoft.Online.SharePoint.PowerShell -DisableNameChecking

# Prompt for credentials
$credentials = Get-Credential

# Connect to SharePoint Online
$adminUrl = "https://addisonprecision-admin.sharepoint.com"
Connect-SPOService -Url $adminUrl -Credential $credentials

# Get all site collections
$siteCollections = Get-SPOSite -Limit All

# Define the user to be added as Site Collection Admin
$user = "migrationwiz@addisonprecision.onmicrosoft.com"

# Add the user as Site Collection Admin to all site collections
foreach ($site in $siteCollections) {
   
    Write-Host $site.Url -ForegroundColor Green
    
    # Add the user as a Site Collection Admin   
    Set-SPOUser -Site $site.Url -LoginName $user -IsSiteCollectionAdmin $true

}

Write-Host "User $user has been added as Site Collection Admin to all SharePoint sites."
