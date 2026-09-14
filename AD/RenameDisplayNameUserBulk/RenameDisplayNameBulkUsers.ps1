# Import active directory module for running AD cmdlets
# Import-Module activedirectory
  
#Store the data from ADUsers.csv in the $ADUsers variable
$Users = Import-csv "PATH TO CSV"

#Loop through each row containing user details in the CSV file 
foreach ($User in $Users)
{
	#Read user data from each field in each row and assign the data to a variable as below
		
	$Username 	= $User.UserPrincipalName
    $SamAccount = $User.SamAccountName
	$Firstname 	= $User.GivenName
	$Lastname 	= $User.SurName
    $DisplayName = $User.DisplayName
    $NewDisplayName = $User.NewDisplayName
    
	#Check to see if the user already exists in AD
	if (Get-ADUser -F {SamAccountName -eq $SamAccount})
	{
		#Account will be updated to the new DisplayName
        Set-ADUser -Identity $SamAccount `
            -DisplayName "$NewDisplayName" `

        #If DisplayName was updated, send confirmation
        Write-Host "A user account with $NewDisplayName has been updated in Active Directory." -ForegroundColor Green
		            
	}
	else
	{
		#User does exist then proceed to create the new user account
		
        #If user does not exist, send warning
        Write-Host "A user account with username $DisplayName does not exist in Active Directory." -ForegroundColor Red

            
	}
}
