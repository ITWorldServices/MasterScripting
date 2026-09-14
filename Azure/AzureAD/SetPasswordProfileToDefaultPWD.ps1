$Password="PlEASEcHANGEmE!2023"
$newPassword = ConvertTo-SecureString $Password -AsPlainText -Force
#Set-MsolUserPassword -UserPrincipalName $userUPN -ForceChangePassword $true
Set-MsolUserPassword -UserPrincipalName $userEmail -NewPassword $newPassword -ForceChangePassword $true -Credential $adminCredential