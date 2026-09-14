#####################################################################################################

###                      Edit the two variables below with your details                           ###


#$Tenant = "schaefges"
$Tenant = Read-Host "Enter Tenant Name. i.e: unitedgroupinc, the part that is found before .onmicrosoft.com"


#$Cred = Get-credential "schaefges@sbigc.com"
$Cred = Get-credential


#####################################################################################################



###   Exchange Online
$cred
$Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://outlook.office365.com/powershell-liveid/ -Credential $cred -Authentication Basic -AllowRedirection
Import-PSSession $Session –AllowClobber


### Exchange Online Protection
$EOPSession = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://ps.protection.outlook.com/powershell-liveid/ -Credential $cred -Authentication Basic -AllowRedirection
Import-PSSession $EOPSession –AllowClobber


### Compliance Center
$ccSession = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri "https://ps.compliance.protection.outlook.com/powershell-liveid/" -Credential $cred -Authentication "Basic" -AllowRedirection
Import-PSSession $ccSession –AllowClobber


### Azure Active Directory Rights Management
Import-Module AADRM
Connect-AadrmService -Credential $cred
    

### Azure Resource Manager
Login-AzureRmAccount -Credential $cred


###   Azure Active Directory v1.0
Import-Module MsOnline
Connect-MsolService -Credential $cred


###  SharePoint Online
Import-Module Microsoft.Online.SharePoint.PowerShell
Connect-SPOService -Url "https://$($Tenant)-admin.sharepoint.com" -Credential $cred


### Skype Online
Import-Module LyncOnlineConnector
Import-Module SkypeOnlineConnector
$SkypeSession = New-CsOnlineSession -Credential $cred
Import-PSSession $SkypeSession 


### Azure AD v2.0
Connect-AzureAD -Credential $cred

