# Monthly Maintenance Office 365 Script
# Last Modified: 10/11/2023


Install-Module -Name MSOnline

# Check if the MSOL service is already connected
if (-not (Get-MsolDomain)) {
    # If not connected, connect to the MSOL service
    Connect-MsolService
}

# Define the output folder
$reportFolder = "C:\Temp\Monthly_Maintenance"

# Create the output folder if it doesn't exist
if (-not (Test-Path -Path $reportFolder)) {
    New-Item -ItemType Directory -Path $reportFolder -Force
}
# Export the list of Global Administrators to a CSV file
$globalAdmins = Get-MsolRoleMember -RoleObjectId $(Get-MsolRole -RoleName "Company Administrator").ObjectId | ForEach-Object {
    if ($_.EmailAddress) {
        $user = Get-MsolUser -UserPrincipalName $_.EmailAddress -ErrorAction SilentlyContinue
        if ($user) {
            # Get the last password change date
            $lastPasswordChangeDate = $user.LastPasswordChangeTimestamp
            [PSCustomObject]@{
                "DisplayName"           = $user.DisplayName
                "EmailAddress"          = $user.UserPrincipalName
                "BlockSignIn"           = $user.BlockCredential
                "LastPasswordChangeDate" = $lastPasswordChangeDate
            }
        } else {
            [PSCustomObject]@{
                "DisplayName"           = $_.DisplayName
                "EmailAddress"          = $_.EmailAddress
                "BlockSignIn"           = "N/A"  # Not available
                "LastPasswordChangeDate" = "N/A"  # Not available
            }
        }
    } else {
        [PSCustomObject]@{
            "DisplayName"           = $_.DisplayName
            "EmailAddress"          = "N/A"  # Not available
            "BlockSignIn"           = "N/A"  # Not available
            "LastPasswordChangeDate" = "N/A"  # Not available
        }
    }
}
# Sort the user data array
$globalAdmins = $globalAdmins | Sort-Object -Property DisplayName, BlockSignIn
# Define the CSV file path for Global Administrators
$globalAdminsCsvFilePath = Join-Path -Path $reportFolder -ChildPath "o365_GlobalAdmins.csv"
# Export the results to a CSV file for Global Administrators
$globalAdmins | Export-Csv -Path $globalAdminsCsvFilePath -NoTypeInformation
Write-Host "Global Admins list exported to $globalAdminsCsvFilePath"



# Define the output folder
#$reportFolder = "C:\Temp\Monthly_Maintenance"
# Create the output folder if it doesn't exist
if (-not (Test-Path -Path $reportFolder)) {
    New-Item -ItemType Directory -Path $reportFolder -Force
}
# Connect to Office 365
#Connect-MsolService
# Get all users excluding guest accounts
$users = Get-MsolUser -All | Where-Object { $_.UserType -ne "Guest" }
# Create an array to store user data
$userData = @()
# Loop through each user
foreach ($user in $users) {
    $userAuthenticationMethod = "N/A"
    $userSignInBlocked = "No"
    $userMFAStatus = "Disabled"

    # Check if the user has MFA enabled
    if ($user.StrongAuthenticationMethods.Count -gt 0) {
        $MFAMethod = $user.StrongAuthenticationMethods | Where-Object { $_.IsDefault -eq $true } | Select-Object -ExpandProperty MethodType
        $Method = ""

        if ($MFAMethod) {
            switch ($MFAMethod) {
                "OneWaySMS" { $Method = "SMS token" }
                "TwoWayVoiceMobile" { $Method = "Phone call verification" }
                "PhoneAppOTP" { $Method = "Hardware token or authenticator app" }
                "PhoneAppNotification" { $Method = "Authenticator app" }
            }
            $userAuthenticationMethod = $Method
            $userMFAStatus = "Enabled"
        }
    }
    
    # Check if the user has MFA enforced
    if ($user.StrongAuthenticationRequirements.Count -gt 0) {
        $userMFAStatus = "Enforced"
    }

    # Check if sign-ins are blocked for the user
    if ($user.BlockCredential) {
        $userSignInBlocked = "Yes"
    }

    $userData += [PSCustomObject]@{
        "SignInBlocked"          = $userSignInBlocked
        "DisplayName"            = $user.DisplayName
        "EmailAddress"           = $user.UserPrincipalName
        "MFAStatus"              = $userMFAStatus
        "AuthenticationMethod"   = $userAuthenticationMethod
    }
}
# Sort the user data array
$userData = $userData | Sort-Object -Property SignInBlocked, MFAStatus, DisplayName
# Define the CSV file path
$csvFilePath = Join-Path -Path $reportFolder -ChildPath "o365_MFAStatus.csv"
# Export the user data to a CSV file
$userData | Export-Csv -Path $csvFilePath -NoTypeInformation
Write-Host "User data exported to $csvFilePath"
