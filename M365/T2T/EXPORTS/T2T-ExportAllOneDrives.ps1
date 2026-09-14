#Set Runtime Parameters
$AdminSiteURL="https://Tenant-admin.sharepoint.com/"

#Connect to SharePoint Online Admin Center
Connect-SPOService -Url $AdminSiteURL 

#Get all OneDrive for Business Site collections
Get-SPOSite -Template "SPSPERS" -Limit ALL -IncludePersonalSite $True | Export-Csv CSVPATH-FILENAME.csv