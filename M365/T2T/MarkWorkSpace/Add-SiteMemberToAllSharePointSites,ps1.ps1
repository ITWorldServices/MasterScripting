# ADD SITE MEMBER to ALL SHAREPOINT SITES
# Load SharePoint Online Management Shell
Import-Module Microsoft.Online.SharePoint.PowerShell -DisableNameChecking

# Prompt for credentials
$credentials = Get-Credential

# Connect to SharePoint Online
$adminUrl = "https://addisonprecision-admin.sharepoint.com"
Connect-SPOService -Url $adminUrl -Credential $credentials

# Get all site collections
$siteCollections = Get-SPOSite -Limit All

# Define the user to be added as Site Member
$user = "migrationwiz@addisonprecision.onmicrosoft.com"

# Add the user as Site Member to all site collections
foreach ($site in $siteCollections) {
    try {
        # Get all groups for the site
        $groups = Get-SPOSiteGroup -Site $site.Url
        
        # Find the Members group
        $membersGroup = $groups | Where-Object { $_.Title -like "*Members*" }
        
        if ($membersGroup) {
            # Add the user to the Members group
            Write-host $membersGroup
            Add-SPOUser -Site $site.Url -Group $membersGroup.Title -LoginName $user
            Write-Host "User $user has been added to the Members group of site $($site.Url)" -ForegroundColor Green
        } else {
            Write-Host "Members group not found for site $($site.Url)" -ForegroundColor Red
        }
    } catch {
        Write-Host "Error processing site $($site.Url): $_" -ForegroundColor Yellow
    }
}

Write-Host "Script execution completed."
