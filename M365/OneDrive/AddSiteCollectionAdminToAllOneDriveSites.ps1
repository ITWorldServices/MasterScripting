#Set Runtime Parameters
$AdminSiteURL="https://Tenant-admin.sharepoint.com/"
$SiteCollAdmin="email address to add permission on all OneDrive sites"
 
#Get Credentials to connect to the SharePoint Admin Center
#$Cred = Get-Credential
 
#Connect to SharePoint Online Admin Center
Connect-SPOService -Url $AdminSiteURL 
 
#Get all OneDrive for Business Site collections
$OneDriveSites = Get-SPOSite -Template "SPSPERS" -Limit ALL -IncludePersonalSite $True
Write-Host -f Yellow "Total Number of OneDrive Sites Found: "$OneDriveSites.count
 
#Add Site Collection Admin to each OneDrive
Foreach($Site in $OneDriveSites)
{
    Write-Host -f Yellow "Adding Site Collection Admin to: "$Site.URL
    Set-SPOUser -Site $Site.Url -LoginName $SiteCollAdmin -IsSiteCollectionAdmin $True
}
Write-Host "Site Collection Admin Added to All OneDrive Sites Successfully!" -f Green