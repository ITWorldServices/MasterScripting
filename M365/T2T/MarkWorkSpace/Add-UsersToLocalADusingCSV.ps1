# Ensure the Active Directory module is available
Import-Module ActiveDirectory

# Path to the CSV file containing users
$users = Import-Csv "C:\temp\Users.csv"

# Specify the target OU (use the distinguishedName format)
$OU = "OU=GSuite Users,OU=Primo Users,DC=Primo,DC=local"

# Loop through each user in the CSV
foreach ($user in $users) {
    New-ADUser -Name $user.Name `
               -GivenName $user.GivenName `
               -Surname $user.Surname `
               -SamAccountName $user.SamAccountName `
               -UserPrincipalName $user.UserPrincipalName `
               -DisplayName $user.DisplayName `
               -EmailAddress $user.EmailAddress `
               -Path $OU `
               -AccountPassword (ConvertTo-SecureString $user.Password -AsPlainText -Force) `
               -ChangePasswordAtLogon $true `
               -Enabled $true
}
