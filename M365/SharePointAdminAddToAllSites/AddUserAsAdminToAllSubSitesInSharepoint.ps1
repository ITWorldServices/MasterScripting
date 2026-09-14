
Commands to add a new user as site administrator to all sub-sites in SharePoint



Install-Module -Name SharePointPnPPowerShellOnline -Force   




**Replace sitename-admin with the company's sharepoint sitename-admin url.
**Replace impactadmin@sitedomain.com with the actual site domain address for 0365

Connect-SPOService -Url https://sitename-admin.sharepoint.com
Get-SPOSite | select url, owner
$SiteUrls=(Get-SPOSite).Url
foreach ($Url in $SiteUrls) {Set-SPOUser -Site $Url -LoginName impactadmin@sitedomain.com -IsSiteCollectionAdmin $true}

