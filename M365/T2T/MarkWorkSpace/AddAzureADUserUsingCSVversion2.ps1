<#
.SYNOPSIS
Automates the addition of new users to Azure Active Directory (Azure AD) from a CSV file, supporting both Standard and GCC High tenants.

.DESCRIPTION
This PowerShell script is designed for IT administrators to efficiently onboard multiple users into Azure Active Directory. It prompts the administrator to select the Azure AD tenant type (Standard or GCC High), connects to the appropriate environment, and imports user details from a specified CSV file. The script validates required fields, checks for existing users, assigns user properties (including manager assignments), and displays all users in an interactive Out-GridView window. It includes robust error handling and supports comprehensive attribute assignment, making it ideal for bulk user creation and migration scenarios.

Key features:
- Tenant selection with color-coded prompts for Standard and GCC High environments.
- Bulk user import from a CSV file with clearly defined columns.
- Validation of required fields (UserPrincipalName and Password).
- Prevention of duplicate user creation by checking for existing accounts.
- Assignment of user attributes such as display name, job title, department, address, phone numbers, and company.
- Optional manager assignment based on display name, with handling for missing or ambiguous managers.
- Interactive display of all Azure AD users using Out-GridView for easy review.
- Clean connection and disconnection from Azure AD.

.SUMMARY
The script streamlines Azure AD user onboarding by automating user creation from CSV data, ensuring data integrity, and providing visual feedback for administrators. It is especially useful for organizations managing Microsoft 365 migrations, onboarding, or tenant transitions where both standard and GCC High environments are in use.
#>

##########################################
##########################################
##                                      ##
## ADD NEW USER TO AZURE AD USING A CSV ##
##                                      ##
##########################################
##########################################

# Clear the screen
Clear-Host

# Function to prompt user for tenant type with colored options
function Get-TenantType {
    Write-Host "Enter tenant type " -NoNewline
    Write-Host "1 for Standard" -ForegroundColor Green -NoNewline
    Write-Host " or " -NoNewline
    Write-Host "2 for GCC High" -ForegroundColor Red
    $tenantType = Read-Host
    switch ($tenantType) {
        "1" { return "Standard" }
        "2" { return "GCCHigh" }
        default { 
            Write-Host "Invalid input. Please enter 1 or 2." -ForegroundColor Red
            return Get-TenantType
        }
    }
}

# Get tenant type from user
$tenantType = Get-TenantType

# Connect to Azure AD based on tenant type
if ($tenantType -eq "Standard") {
    Connect-AzureAD
} else {
    Connect-AzureAD -AzureEnvironmentName AzureUSGovernment
}

##########################################################################################
# CSV FILE REQUIREMENTS                                                                  #
# CSV columns: UserPrincipalName, DisplayName, GivenName, Surname, JobTitle, Department, #
# StreetAddress, City, State, Country, PostalCode, MobilePhone, TelephoneNumber,         #
# ManagerDisplayName, Company, Password                                                  #
##########################################################################################
$csvPath = "C:\temp\Chandler-Mailboxes.csv"
$usersToAdd = Import-Csv -Path $csvPath

foreach ($user in $usersToAdd) {
    # Defensive: skip user if required fields are missing
    if ([string]::IsNullOrWhiteSpace($user.UserPrincipalName)) {
        Write-Host "Skipping row: UserPrincipalName is missing." -ForegroundColor Red
        continue
    }
    if ([string]::IsNullOrWhiteSpace($user.Password)) {
        Write-Host "Skipping $($user.UserPrincipalName): Password is missing." -ForegroundColor Red
        continue
    }

    $params = @{
        UserPrincipalName = $user.UserPrincipalName
        AccountEnabled    = $true
        MailNickName      = ($user.UserPrincipalName -split '@')[0]
    }
    if (-not [string]::IsNullOrWhiteSpace($user.DisplayName))      { $params.DisplayName      = $user.DisplayName }
    if (-not [string]::IsNullOrWhiteSpace($user.GivenName))        { $params.GivenName        = $user.GivenName }
    if (-not [string]::IsNullOrWhiteSpace($user.Surname))          { $params.Surname          = $user.Surname }
    if (-not [string]::IsNullOrWhiteSpace($user.JobTitle))         { $params.JobTitle         = $user.JobTitle }
    if (-not [string]::IsNullOrWhiteSpace($user.Department))       { $params.Department       = $user.Department }
    if (-not [string]::IsNullOrWhiteSpace($user.StreetAddress))    { $params.StreetAddress    = $user.StreetAddress }
    if (-not [string]::IsNullOrWhiteSpace($user.City))             { $params.City             = $user.City }
    if (-not [string]::IsNullOrWhiteSpace($user.State))            { $params.State            = $user.State }
    if (-not [string]::IsNullOrWhiteSpace($user.PostalCode))       { $params.PostalCode       = $user.PostalCode }
    if (-not [string]::IsNullOrWhiteSpace($user.Country))          { $params.Country          = $user.Country }
    if (-not [string]::IsNullOrWhiteSpace($user.TelephoneNumber))  { $params.TelephoneNumber  = $user.TelephoneNumber }
    if (-not [string]::IsNullOrWhiteSpace($user.MobilePhone))      { $params.Mobile           = $user.MobilePhone }
    if (-not [string]::IsNullOrWhiteSpace($user.Company))          { $params.CompanyName      = $user.Company }

    # Handle password profile
    $PasswordProfile = New-Object -TypeName Microsoft.Open.AzureAD.Model.PasswordProfile
    $PasswordProfile.Password = $user.Password
    $params.PasswordProfile = $PasswordProfile

    # Check if the user already exists
    $existingUser = Get-AzureADUser -Filter "UserPrincipalName eq '$($user.UserPrincipalName)'"
    if ($null -eq $existingUser) {
        $newUser = New-AzureADUser @params
        Write-Host "New user added: $($user.UserPrincipalName)" -ForegroundColor Green

        # Set the manager if provided
        if (-not [string]::IsNullOrWhiteSpace($user.ManagerDisplayName)) {
            $manager = Get-AzureADUser -Filter "DisplayName eq '$($user.ManagerDisplayName)'"
            if ($null -ne $manager) {
                if ($manager.Count -gt 1) {
                    Write-Host "Multiple managers found with display name '$($user.ManagerDisplayName)' for user $($user.UserPrincipalName). Manager not set."
                } else {
                    Set-AzureADUserManager -ObjectId $newUser.ObjectId -RefObjectId $manager.ObjectId
                    Write-Host "Manager set for $($user.UserPrincipalName)"
                }
            } else {
                Write-Host "Manager not found for $($user.UserPrincipalName)"
            }
        }
    } else {
        Write-Host "User already exists: $($user.UserPrincipalName)" -ForegroundColor Red
    }
}

# Get all Azure AD users and display in Out-GridView
$allUsers = Get-AzureADUser -All $true | Select-Object UserPrincipalName, DisplayName, Department, JobTitle
$allUsers | Out-GridView -Title "All Azure AD Users"

Disconnect-AzureAD
