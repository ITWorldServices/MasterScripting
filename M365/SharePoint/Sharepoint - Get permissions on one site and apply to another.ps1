#First yoou have hopefully copied everything over to the new location
#Enter your correct domain in th line below
Connect-SPOService -Url https://brehobcorp-admin.sharepoint.com

#Enter you correct site in the line below
$sourceLibraryUrl = "https://brehobcorp.sharepoint.com/sites/allcompany/Shared Documents"
$permissions = Get-SPOSiteGroup -Site $sourceLibraryUrl
#Change the path if you do not want it stored in C:\temp
$permissions | Export-Csv -Path "C:\temp\sppermissions.csv"

#Input the place to import the permissions to
$targetLibraryUrl = "https://brehobcorp.sharepoint.com/sites/ImpactTestSP/Shared Documents/01 stockroom"
$permissions = Import-Csv -Path "C:\temp\sppermissions.csv"
foreach ($permission in $permissions) {
    Add-SPOUser -Site $targetLibraryUrl -LoginName $permission.LoginName -Group $permission.Group
}