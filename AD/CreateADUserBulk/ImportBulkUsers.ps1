# Import active directory module for running AD cmdlets
Import-Module activedirectory
  
#Store the data from ADUsers.csv in the $ADUsers variable
$Users = Import-csv "PATH TO CSV FILE"

#Loop through each row containing user details in the CSV file 
foreach ($User in $Users)
{
	#Read user data from each field in each row and assign the data to a variable as below
		
	$Username 	= $User.UserPrincipalName
    $SamAccount = $User.SamAccountName
	$Password 	= $User.Password
	$Firstname 	= $User.FirstName
	$Lastname 	= $User.LastName
    $Displayname = $User.DisplayName
    $Mail = $User.ProxyAddresses
    $Mobile = $User.PhoneNumber
    $OU = $User.OU #This field refers to the OU the user account is to be created in

	#Check to see if the user already exists in AD
	if (Get-ADUser -F {SamAccountName -eq $Username})
	{
		 #If user does exist, give a warning
		 Write-Warning "A user account with username $Username already exist in Active Directory."
	}
	else
	{
		#User does not exist then proceed to create the new user account
		
        #Account will be created in the OU provided by the $OU variable read from the CSV file
		New-ADUser `
            -EmailAddress $Mail `
            -SamAccountName $SamAccount `
            -UserPrincipalName $Username `
            -MobilePhone $Mobile `
            -Name "$Firstname $Lastname" `
            -GivenName $Firstname `
            -Surname $Lastname `
            -Enabled $True `
            -DisplayName "$Displayname" `
            -Path $OU `
            -AccountPassword (convertto-securestring $Password -AsPlainText -Force) 
	}
}
