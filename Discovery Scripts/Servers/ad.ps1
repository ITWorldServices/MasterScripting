# Monthly Maintenance Active Directory Script
# Last Modified: 10/3/2023


$reportFolder = "C:\Temp\Monthly_Maintenance"
New-Item -Path $reportFolder -ItemType Directory -Force | Out-Null

### List of all the disabled users
# Define the output CSV file path
$csvFilePath = "$reportFolder\DisabledUsers.csv"
# Get disabled users and their OUs
$disabledUsers = Get-ADUser -Filter {Enabled -eq $false} -Properties DistinguishedName,Enabled,Name,SamAccountName
$usersInfo = @()
# Get their OU locations
foreach ($user in $disabledUsers) {
    $userInfo = [PSCustomObject]@{
        AccountActive = $user.Enabled
        UserName = $user.SamAccountName
        DisplayName = $user.Name
        OU = ($user.DistinguishedName -split ",", 2)[1] -replace "OU=", ""
    }
    $usersInfo += $userInfo
}
# Display the list of user
#$userInfo | Format-Table -AutoSize AccountActive, UserName, DisplayName, OU
# Export to CSV
$usersInfo | Export-Csv -Path $csvFilePath -NoTypeInformation
# Echo file location
Write-Host "Disabled users list and their OUs have been exported to $csvFilePath"



### List of all users in evaluated security groups
# Define the list of security group names
$groups = @("Domain Admins", "Enterprise Admins", "Schema Admins", "Administrators")
# Loop through the list of security groups
foreach ($group in $groups) {
    $csvFilePath = "$reportFolder\Admins_$group.csv"
    $csvFilePath = $csvFilePath -replace ' ', ''
    # Import the Active Directory module
    Import-Module ActiveDirectory
    # Get the members of the security group
    $securityGroup = Get-ADGroup $group
    $securityGroupMembers = Get-ADGroupMember $securityGroup -Recursive
    # Create an array to store user information
    $usersInfo = @()
    # Loop through the members and gather user information
    foreach ($member in $securityGroupMembers) {
        if ($member.objectClass -eq "user") {
            $user = Get-ADUser $member -Properties DisplayName, SamAccountName, UserPrincipalName, Enabled
            $userInfo = [PSCustomObject]@{
                SecurityGroup = $group
                UserName = $user.SamAccountName
                DisplayName = $user.Name
                AccountActive = $user.Enabled
            }
            $usersInfo += $userInfo
        }
    }
    # Display the list of user
    #$userInfo | Format-Table -AutoSize SecurityGroup, UserName, DisplayName, AccountActive
    # Export the user information to a CSV file
    $usersInfo | Export-Csv -Path $csvFilePath -NoTypeInformation
    Write-Host "$group users list have been exported to $csvFilePath"
}



### Inactive Users
# Define the path to the CSV file
$csvFilePath = "$reportFolder\InactiveUsers.csv"
# Define the number of days of inactivity
$inactiveDays = 180
# Get the current date
$currentDate = Get-Date
# Calculate the date 180 days ago
$lastActiveDate = $currentDate.AddDays(-$inactiveDays)
# Get a list of enabled user accounts that have been inactive for at least 180 days
$inactiveUsers = Get-ADUser -Filter { (Enabled -eq $true) -and (LastLogonTimestamp -lt $lastActiveDate) } -Properties LastLogonTimestamp
# Define an array to store the results
$usersInfo = @()
# Process each user to create a custom object with the desired properties
foreach ($user in $inactiveUsers) {
    $userInfo = [PSCustomObject]@{
        UserName = $user.SamAccountName
        DisplayName = $user.Name
        AccountActive = $user.Enabled
        LastLogonDate = [DateTime]::FromFileTime($user.LastLogonTimestamp).ToString('MM/dd/yyyy')
        OU = ($user.DistinguishedName -split ",", 2)[1] -replace "OU=", ""
    }
    $usersInfo += $userInfo
}
# Display the list of user
#$userInfo | Format-Table -AutoSize UserName, DisplayName, AccountActive, LastLogonDate
# Export the list of inactive users to a CSV file
$usersInfo | Export-Csv -Path $csvFilePath -NoTypeInformation
# Write a message to the console indicating where the CSV file has been exported
Write-Host "Inactive users list have been exported to $csvFilePath"



### List User Accounts with Passwords Set to Never Expire
# Define the path to the CSV file
$csvFilePath = "$reportFolder\NeverExpireUserPasswords.csv"
# Get user accounts with passwords set to never expire
$neverExpireUsers = Get-ADUser -Filter {PasswordNeverExpires -eq $true} -Properties Name, SamAccountName, Enabled, DisplayName
# Create an array to store user information
$userInfo = @()
# Process each user and create a custom object with the desired properties
foreach ($user in $neverExpireUsers) {
    $userProperties = [PSCustomObject]@{
        UserName = $user.SamAccountName
        DisplayName = $user.Name
        AccountActive = $user.Enabled
        PasswordNeverExpires = $true
#        OU = ($user.DistinguishedName -split ",", 2)[1] -replace "OU=", ""
    }
    $userInfo += $userProperties
}
# Display the list of user
#$userInfo | Format-Table -AutoSize UserName, DisplayName, AccountActive, PasswordNeverExpires
# Export the list of users to a CSV file
$userInfo | Export-Csv -Path $csvFilePath -NoTypeInformation
Write-Host "List of users with passwords set to never expire exported to $csvFilePath"



### List User Accounts with no passwords set
# Define the path to the CSV file
$csvFilePath = "$reportFolder\UsersWithNoPasswords.csv"
# Search for user accounts with no passwords (empty passwords)
$usersWithNoPasswords = Get-ADUser -Filter {PasswordLastSet -eq 0} -Properties Name, SamAccountName, Enabled
# Create an array to store user information
$userInfo = @()
# Process each user and create a custom object with the desired properties
foreach ($user in $usersWithNoPasswords) {
    $userInfoEntry = [PSCustomObject]@{
        UserName = $user.SamAccountName
        DisplayName = $user.Name
        AccountActive = $user.Enabled
        PasswordBlank = $true
#        OU = ($user.DistinguishedName -split ",", 2)[1] -replace "OU=", ""
}
    $userInfo += $userInfoEntry
}
# Display the list of user accounts with no passwords
#$userInfo | Format-Table -AutoSize UserName, DisplayName, AccountActive, PasswordBlank
# Export the list of user accounts with no passwords to a CSV file
$userInfo | Export-Csv -Path $csvFilePath -NoTypeInformation
Write-Host "List of user accounts with no passwords exported to $csvFilePath"



### Find Computers Inactive for 180 Days
# Define the path to the CSV file
$csvFilePath = "$reportFolder\InactiveComputers.csv"
# Define the number of days of inactivity
$inactiveDays = 180
# Get the current date
$currentDate = Get-Date
# Calculate the date 180 days ago
$lastActiveDate = $currentDate.AddDays(-$inactiveDays)
# Search for computer accounts that have been inactive for at least 180 days
$inactiveComputers = Get-ADComputer -Filter {LastLogonDate -lt $lastActiveDate} -Properties Name, LastLogonDate, DistinguishedName
# Create an array to store computer information
$computerInfo = @()
# Process each inactive computer and gather information
foreach ($computer in $inactiveComputers) {
    # Get the OU location by extracting it from the DistinguishedName
    $ou = ($computer.DistinguishedName -split ",", 2)[1] -replace "OU=", ""
    $computerInfoEntry = [PSCustomObject]@{
        Name = $computer.Name
        LastLogonDate = $computer.LastLogonDate
        OU = $ou
    }
    $computerInfo += $computerInfoEntry
}
# Display the list of inactive computers
#$computerInfo | Format-Table -AutoSize Name, LastLogonDate, OU
# Export the list of inactive computers to a CSV file
$computerInfo | Export-Csv -Path $csvFilePath -NoTypeInformation
Write-Host "List of inactive computers exported to $csvFilePath"






