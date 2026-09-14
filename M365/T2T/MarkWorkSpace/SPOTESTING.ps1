
Connect-SPOService -Url https://addisonprecision-admin.sharepoint.com

#Get-SPOSite -Limit All | Select Url, Title, Owner

Get-SPOSite -Limit All | Select *

Get-SPOSite -Limit All | Select Title 

#Get-SPOSite -Limit All | Get-Member

#Get-SPOSite -Identity <site-url> | Select *
