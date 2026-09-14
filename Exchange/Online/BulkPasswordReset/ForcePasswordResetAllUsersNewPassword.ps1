Import-Csv "INSERT PATH TO GET-MsolUsers -All exported CSV HERE Be sure you added the NewPassword column and defined your passwords" | ForEach-Object {
$upn = $_.UserPrincipalName
$newpassword = $_.NewPassword
Set-MsolUserPassword -UserPrincipalName $upn -ForceChangePasswordOnly $true -ForceChangePassword $true;
}