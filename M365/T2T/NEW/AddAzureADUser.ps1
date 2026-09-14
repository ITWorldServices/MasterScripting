$PasswordProfile=New-Object -TypeName Microsoft.Open.AzureAD.Model.PasswordProfile
$PasswordProfile.Password="MJfmz23^Da6Ac2n!Eu5UKY"
New-AzureADUser -DisplayName "David Busch" -GivenName "David" -SurName "Busch" -UserPrincipalName david.busch@icsholdingllc.onmicrosoft.com -UsageLocation US -MailNickName david.busch -PasswordProfile $PasswordProfile -AccountEnabled $true

$PasswordProfile=New-Object -TypeName Microsoft.Open.AzureAD.Model.PasswordProfile
$PasswordProfile.Password="<user account password>"
New-AzureADUser -DisplayName "<display name>" -GivenName "<first name>" -SurName "<last name>" -UserPrincipalName <sign-in name> -UsageLocation <ISO 3166-1 alpha-2 country code> -MailNickName <mailbox name> -PasswordProfile $PasswordProfile -AccountEnabled $true