Import-Csv "PATH TO CSV FILE" | ForEach-Object {
$upn = $_.UserPrincipalName
$newupn = $_.NewUserPrincipalName
Set-MsolUserPrincipalName -UserPrincipalName $upn -NewUserPrincipalName $newupn;
}