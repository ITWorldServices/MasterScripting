# Synopsis:
# This PowerShell script connects to Azure Active Directory (Azure AD) and resets passwords for all non-admin users.
# It prompts the user to confirm the operation and allows customization of the new password's length and character set.
# The script exports a CSV file containing the users and their new passwords.

# Details:
# - Requires the AzureAD PowerShell module.
# - Prompts for confirmation before resetting passwords.
# - Generates strong, random passwords based on user input for length and character types.
# - Exports user data to a CSV file for record-keeping.

# Description:
# This script is designed to automate the process of resetting passwords for non-admin users in Azure Active Directory.
# It ensures that admin roles are excluded from the password reset operation, providing a secure and controlled environment for managing user accounts.
# The script includes interactive prompts for user input, ensuring that the administrator is aware of the actions being performed.

# Clear the screen
Clear-Host

# Function to get yes/no input from the user
function Get-YesNoInput {
    param (
        [string]$Prompt
    )
    do {
        $inputValue = Read-Host "$Prompt (y/n)"
        $inputValue = $inputValue.Trim().ToLower()
        if ($inputValue -notin @('y', 'n')) {
            Write-Host "Invalid input. Please enter y for yes or n for no." -ForegroundColor Red
        }
    } until ($inputValue -in @('y', 'n'))
    return $inputValue -eq 'y'
}

# Prompt the user to connect to Azure AD and reset passwords
$ResetPasswords = Get-YesNoInput "Do you want to connect to Azure AD and reset passwords for non-admin users?" -ForegroundColor Yellow

if ($ResetPasswords) {
    # Check if AzureAD module is installed
    if (-not (Get-Module -ListAvailable AzureAD)) {
        Write-Host "AzureAD module is not installed. Installing..." -ForegroundColor Yellow
        Install-Module -Name AzureAD -Force
    }

    # Connect to Azure AD on GCC HIGH
    try {
        Connect-AzureAD -AzureEnvironmentName AzureUSGovernment
    } catch {
        Write-Host "Error connecting to Azure AD: $($Error[0].Message)" -ForegroundColor Red
        exit
    }

    # Get the tenant name
    $tenantDetail = Get-AzureADTenantDetail
    $tenantName = $tenantDetail.DisplayName

    # Define all Entra ID admin roles
    $adminRoles = @(
        "Company Administrator", 
        "Global Administrator", 
        "Privileged Role Administrator",
        "User Administrator",
        "Application Administrator",
        "Groups Administrator",
        "Helpdesk Administrator",
        "Hybrid Identity Administrator",
        "License Administrator",
        "SharePoint Administrator",
        "Intune Administrator",
        "Exchange Administrator",
        "Dynamics 365 Administrator",
        "Compliance Administrator",
        "Billing Administrator",
        "Aplication Administrator",
        "Privileged Authentication Administrator",
        "Reports Reader",
        "Authentication Administrator",
        "B2C IEF Keyset Administrator",
        "B2C IEF Policy Administrator",
        "Cloud Application Administrator",
        "Cloud Device Administrator",
        "Conditional Access Administrator",
        "Directory Readers",
        "Directory Synchronization Accounts",
        "Directory Writers",
        "External ID User Flow Administrator",
        "External ID User Flow Attribute Administrator",
        "External Identity Provider Administrator",
        "Guest Inviter",
        "Partner Tier1 Support",
        "Partner Tier2 Support",
        "People Administrator", # New role
        "IoT Device Administrator" # New role
    )

    # Get all users
    $allUsers = Get-AzureADUser -All $true

    # Get users with admin roles
    $adminUsers = $allUsers | Where-Object {
        $userRoles = Get-AzureADUserMembership -ObjectId $_.ObjectId
        $userRoles | Where-Object { $_.DisplayName -in $adminRoles }
    }

    # Get users without admin roles
    $usersToReset = $allUsers | Where-Object {
        $userRoles = Get-AzureADUserMembership -ObjectId $_.ObjectId
        -not ($userRoles | Where-Object { $_.DisplayName -in $adminRoles })
    }

    # Display users whose passwords will be changed
    Write-Host "The following users will have their passwords reset:" -ForegroundColor Yellow
    foreach ($user in $usersToReset) {
        Write-Host "  - $($user.UserPrincipalName)" -ForegroundColor Cyan
    }

    # Display admins whose passwords will not be changed
    Write-Host # This adds a blank line
    Write-Host "The following admins will NOT have their passwords reset:" -ForegroundColor Yellow
    foreach ($admin in $adminUsers) {
        Write-Host "  - $($admin.UserPrincipalName)" -ForegroundColor Red
    }

    # Prompt for confirmation before proceeding
    Write-Host # This adds a blank line
    $ConfirmReset = Get-YesNoInput "Are you sure you want to proceed with resetting passwords for these users?"
    if (-not $ConfirmReset) {
        Write-Host "Password reset operation cancelled." -ForegroundColor Red
        exit
    }

    # Function to generate a strong password
    function Generate-StrongPassword {
        param (
            [int]$Length,
            [bool]$IncludeUppercase,
            [bool]$IncludeLowercase,
            [bool]$IncludeNumbers,
            [bool]$IncludeSpecialChars
        )

        $chars = @()
        if ($IncludeUppercase) { $chars += [char[]]"ABCDEFGHJKLMNPQRSTUVWXYZ" }
        if ($IncludeLowercase) { $chars += [char[]]"abcdefghijkmnpqrstuvwxyz" }
        if ($IncludeNumbers) { $chars += [char[]]"23456789" }
        if ($IncludeSpecialChars) { $chars += [char[]]"!#$%^&-;:<>?" }

        # Ensure password includes at least one of each selected character type
        $password = @()
        if ($IncludeUppercase) { $password += Get-Random -InputObject ([char[]]"ABCDEFGHJKLMNPQRSTUVWXYZ") }
        if ($IncludeLowercase) { $password += Get-Random -InputObject ([char[]]"abcdefghijkmnpqrstuvwxyz") }
        if ($IncludeNumbers) { $password += Get-Random -InputObject ([char[]]"23456789") }
        if ($IncludeSpecialChars) { $password += Get-Random -InputObject ([char[]]"!#$%^&-;:<>?") }

        # Fill the rest of the password length with random characters from the selected sets
        for ($i = $password.Count; $i -lt $Length; $i++) {
            $password += Get-Random -InputObject $chars
        }

        # Shuffle the array to avoid the first characters always being from specific sets
        $password = $password | Get-Random -Count $password.Count

        -join $password
    }

    # Prompt for password length and character set options
    do {
        $PasswordLength = Read-Host "Enter the length of the passwords (default is 16)"
        if (-not [int]::TryParse($PasswordLength, [ref]$null)) {
            $PasswordLength = 16
            Write-Host "Invalid input. Using default length of 16." -ForegroundColor Yellow
        } elseif ([int]$PasswordLength -le 0) {
            Write-Host "Invalid input. Please enter a positive integer." -ForegroundColor Red
        }
    } until ([int]::TryParse($PasswordLength, [ref]$null) -and [int]$PasswordLength -gt 0)
    $PasswordLength = [int]$PasswordLength

    $IncludeUppercase = Get-YesNoInput "Include uppercase letters?"
    $IncludeLowercase = Get-YesNoInput "Include lowercase letters?"
    $IncludeNumbers = Get-YesNoInput "Include numbers?"
    $IncludeSpecialChars = Get-YesNoInput "Include special characters?"
    Write-Host # This adds a blank line

    # Ensure at least one character type is selected
    while (-not ($IncludeUppercase -or $IncludeLowercase -or $IncludeNumbers -or $IncludeSpecialChars)) {
        Write-Host "Please select at least one character type." -ForegroundColor Red
        $IncludeUppercase = Get-YesNoInput "Include uppercase letters?"
        $IncludeLowercase = Get-YesNoInput "Include lowercase letters?"
        $IncludeNumbers = Get-YesNoInput "Include numbers?"
        $IncludeSpecialChars = Get-YesNoInput "Include special characters?"
    }

    # Create an array to store the users and their new passwords
    $resetPasswords = @()

    # Reset passwords for non-admin users
    foreach ($user in $usersToReset) {
        $newPassword = Generate-StrongPassword -Length $PasswordLength -IncludeUppercase $IncludeUppercase -IncludeLowercase $IncludeLowercase -IncludeNumbers $IncludeNumbers -IncludeSpecialChars $IncludeSpecialChars
        try {
            Set-AzureADUserPassword -ObjectId $user.ObjectId -Password (ConvertTo-SecureString -String $newPassword -AsPlainText -Force)
            Write-Host "Password reset for user $($user.UserPrincipalName) with new password: $newPassword" -ForegroundColor Green
            $resetPasswords += [PSCustomObject]@{
                UserPrincipalName = $user.UserPrincipalName
                NewPassword       = $newPassword
            }
        } catch {
            Write-Host "Error resetting password for user $($user.UserPrincipalName): $($Error[0].Message)" -ForegroundColor Red
        }
    }

    # Export the users and their new passwords to a CSV file
    if ($resetPasswords.Count -gt 0) {
        $CurrentDateTime = Get-Date -Format "yyyyMMdd_HHmmss"
        $fileName = "C:\temp\$($tenantName)_ResetPasswords_$CurrentDateTime.csv"
        $resetPasswords | Export-Csv -Path $fileName -NoTypeInformation -Encoding UTF8
        Write-Host # This adds a blank line
        Write-Host "Users and their new passwords exported to: $fileName" -ForegroundColor Yellow
        Write-Host # This adds a blank line
    }
}
