#Variables for processing
$AdminURL = "https://firsthospitality-admin.sharepoint.com/"
$AdminName = "aitadmin@firsthospitality.com"
$SearchUser = "lparisi@firsthospitality.com"
  
#User Names Password to connect
$Password = Read-host -assecurestring "Enter Password for $AdminName"
$Credential = new-object -typename System.Management.Automation.PSCredential -argumentlist $AdminName, $Password
 
Try {
    #Connect to SharePoint Online
    Connect-SPOService -url $AdminURL -credential $Credential
 
    #Get all Site colections
    $Sites = Get-SPOSite -Limit ALL
 
    Foreach ($Site in $Sites)
    {
        Write-host $Site.URL
     
        #Get all Site Collection Administrators
        $SiteAdmins = Get-SPOUser -Site $Site.Url -Limit ALL | Where { $_.IsSiteAdmin -eq $True}
        foreach($Admin in $SiteAdmins)
        {
            Write-host $Admin.LoginName
            if ($Admin.LoginName -eq $SearchUser)
            {
                Write-host -BackgroundColor Red -ForegroundColor White "Removing Admin" $SearchUser "from" $Site.Url "SharePoint Site"
                Set-SPOUser -LoginName $SearchUser -IsSiteCollectionAdmin $false -Site $Site.Url
            }      
        }
    }
}
catch {
    write-host "Error: $($_.Exception.Message)" -foregroundcolor Red
}