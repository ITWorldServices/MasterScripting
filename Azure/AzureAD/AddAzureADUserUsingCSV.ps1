##########################################
##########################################
##                                      ##
## ADD NEW USER TO AZURE AD USING A CSV ##
##                                      ##
##########################################
##########################################

# TO connect to a Standard Tenant
# Connect-AzureAD
 
# TO connect to a GCC High Tenant
Connect-AzureAD -AzureEnvironmentName AzureUSGovernment
 
# CSV FILE REQUIREMENTS
# CSV columns: UserPrincipalName, DisplayName, GivenName, Surname, JobTitle, Department, 
# StreetAddress, City, State, Country, PostalCode, MobilePhone, TelephoneNumber, 
# ManagerDisplayName, Company, Password
$csvPath = "C:\temp\NewUsers.csv"
$usersToAdd = Import-Csv -Path $csvPath

 
foreach ($user in $usersToAdd) {
    $userPrincipalName = $user.UserPrincipalName
    $displayName = $user.DisplayName
    $givenName = $user.GivenName
    $surname = $user.Surname
    $jobTitle = $user.JobTitle
    $department = $user.Department
    $streetAddress = $user.StreetAddress
    $city = $user.City
    $state = $user.State
    $country = $user.Country
    $postalCode = $user.PostalCode
    $mobilePhone = $user.MobilePhone
    $telephoneNumber = $user.TelephoneNumber
    $managerDisplayName = $user.ManagerDisplayName
    $companyName = $user.Company
    $password = $user.Password
 
    # Check if the user already exists
    $existingUser = Get-AzureADUser -Filter "UserPrincipalName eq '$userPrincipalName'"
 
    if ($null -eq $existingUser) {
        # Create a password profile
        $PasswordProfile = New-Object -TypeName Microsoft.Open.AzureAD.Model.PasswordProfile
        $PasswordProfile.Password = $password
 
        # Create the new user with additional attributes
        $newUser = New-AzureADUser -UserPrincipalName $userPrincipalName `
                        -DisplayName $displayName `
                        -GivenName $givenName `
                        -Surname $surname `
                        -JobTitle $jobTitle `
                        -Department $department `
                        -StreetAddress $streetAddress `
                        -City $city `
                        -State $state `
                        -PostalCode $postalCode `
                        -Country $country `
                        -TelephoneNumber $telephoneNumber `
                        -Mobile $mobilePhone `
                        -CompanyName $companyName `
                        -PasswordProfile $PasswordProfile `
                        -AccountEnabled $true `
                        -MailNickName ($userPrincipalName -split '@')[0]
 
        Write-Host "New user added: $userPrincipalName"
 
        # Set the manager if provided
        if (-not [string]::IsNullOrEmpty($managerDisplayName)) {
            $manager = Get-AzureADUser -Filter "DisplayName eq '$managerDisplayName'"
            if ($null -ne $manager) {
                if ($manager.Count -gt 1) {
                    Write-Host "Multiple managers found with display name '$managerDisplayName' for user $userPrincipalName. Manager not set."
                } else {
                    Set-AzureADUserManager -ObjectId $newUser.ObjectId -RefObjectId $manager.ObjectId
                    Write-Host "Manager set for $userPrincipalName"
                }
            } else {
                Write-Host "Manager not found for $userPrincipalName"
                Write-Host 
            }
        }
    } else {
        Write-Host "User already exists: $userPrincipalName"
    }
}
 
Disconnect-AzureAD