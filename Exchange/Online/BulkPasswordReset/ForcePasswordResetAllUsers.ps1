Import-Csv "EXPORT MSOL USERS TO CSV AND INSERT PATH TO CSV FILE HERE" | ForEach-Object {
$upn = $_.UserPrincipalName
Set-MsolUserPassword -UserPrincipalName $upn -ForceChangePasswordOnly $true -ForceChangePassword $true;
}