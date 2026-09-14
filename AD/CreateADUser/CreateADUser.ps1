# Import required modules
Import-Module ActiveDirectory

# Create a new password that is not prompted for
# $securepassword = ConvertTo-SecureString "R0@dBu1Lder" -AsPlainText -Force

#Prompt for default user password which will be applied to all created users
$securepassword = Read-Host -Prompt "Please enter the default user password to apply to all new users"

# Prompt user for the CVS file path
$filepath = Read-Host -Prompt "Please enter the path to your CSV file"

# Import the file into a variable
$users = Import-Csv $filepath

# Loop through each row and gather information
ForEach ($user in $users) {

    # Gather the user's information
    $fname = $user.FirstName
    $lname = $user.LastName
    $userprincipalname = $user.UserPrincipalName
    $jobtitle = $user.JobTitle
    $streetaddress = $user.Address
    $city = $user.City
    $state = $user.State
    $postalcode = $user.PostalCode
    $officephone = $user.OfficePhone
    $mobilephone = $user.MobilePhone
    $fax = $user.Fax
    $displayname = $user.DisplayName
    $emailaddress = $user.EmailAddress
  

    # Create new user for each user in CSV file
    New-ADUser -Name "$fname $lname" -GivenName $fname -Surname $lname -UserPrincipalName $userprincipalname -DisplayName $displayname -Title $jobtitle -StreetAddress $streetaddress -City $city -State $state -PostalCode $postalcode -OfficePhone $officephone -MobilePhone $mobilephone -Fax $fax -EmailAddress $emailaddress -AccountPassword $securepassword -ChangePasswordAtLogon $False -Enabled $True -OtherAttributes @{'proxyAddresses'= $proxyAddresses}

    # Echo output for each new user
    echo "Account created for $fname $lname"
}