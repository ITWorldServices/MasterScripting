Connect-SPOService -Url "https://primocenter-admin.sharepoint.com"


# Syntax is email address
Request-SPOPersonalSite -UserEmails $userEmail

# Wait Until Site is Present. Verify using SharePoint Admin Center in the "More" -> User Profiles -> Manage Profiles
# Syntax for the users are username_domainname_com
Set-SPOUser -Site https://primocenter-my.sharepoint.com/personal/$userEmail -LoginName "migrationwiz@primocenter.onmicrosoft.com" -IsSiteCollectionAdmin $true