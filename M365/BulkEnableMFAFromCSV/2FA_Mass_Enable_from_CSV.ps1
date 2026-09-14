Import-CSv -Path "users.csv" | ForEach {
$UPN=$_.UserPrincipalName
$Users=Get-MsolUser -User $UPN
$Auth= New-Object -TypeName Microsoft.Online.Administration.StrongAuthenticationRequirement
$Auth.RelyingParty = “*”
$MFA= @($Auth)

#Enable MFA for a user
#Set-MsolUser -UserPrincipalName test1@domain.com -StrongAuthenticationRequirements $mfa
 
#Enable MFA for all users in the CSV files(use with CAUTION!)
#You have to create app password after strong auth configuration

$Users | Set-MsolUser -StrongAuthenticationRequirements $MFA

#Disable strong auth requirement
# $Users | Set-MsolUser -StrongAuthenticationRequirements $False
}
