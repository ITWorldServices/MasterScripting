<#
.SYNOPSIS
This script adds users to specified security groups in Azure AD using Microsoft Graph.

.DESCRIPTION
This PowerShell script performs the following tasks:
1. Asks the user to specify the tenant type (standard or GCC High).
2. Prompts for the path to a CSV file containing user principal names and corresponding security group names.
   - A default path of C:\Temp\Users.csv is provided.
   - If no input is given, the default path is used.
   - If the file doesn't exist, it will continue to prompt until a valid file path is provided.
3. Presents a verification of the chosen file path before proceeding.
4. Validates the CSV file to ensure it contains the required columns: userPrincipalName and SecurityGroup.
5. For each user in the CSV:
   a. Checks if the specified security group exists.
   b. Verifies if the user exists.
   c. If both user and group exist, checks if the user is already a member of the group.
   d. If the user is not a member, adds the user to the group.
6. All actions and results are logged to both the console and a CSV output file.
7. The output file is named with the tenant name and a timestamp for easy identification.
8. Displays a summary of the operations performed, including total users processed, successful additions, users already in groups, and errors encountered.

.NOTES
- Requires the Microsoft.Graph.Users and Microsoft.Graph.Groups modules.
- Requires appropriate permissions to read and modify users and groups in Azure AD.
- Input CSV must have columns: userPrincipalName, SecurityGroup
- The script includes error handling and will provide feedback for various error conditions.
- Default input file path is C:\Temp\Users.csv
#>

# Error handling function
function Write-ErrorLog {
    param (
        [string]$Message,
        [string]$ErrorDetails
    )
    Write-Host "Error: $Message" -ForegroundColor Red
    if ($ErrorDetails) {
        Write-Host "Details: $ErrorDetails" -ForegroundColor Red
    }
}

# Input validation function
function Test-InputFile {
    param (
        [string]$FilePath
    )
    if (-not (Test-Path $FilePath)) {
        Write-ErrorLog "Input file not found: $FilePath"
        return $false
    }
    $csvContent = Import-Csv -Path $FilePath
    if (-not ($csvContent | Get-Member -Name "userPrincipalName" -MemberType NoteProperty) -or
        -not ($csvContent | Get-Member -Name "SecurityGroup" -MemberType NoteProperty)) {
        Write-ErrorLog "CSV file must contain 'userPrincipalName' and 'SecurityGroup' columns"
        return $false
    }
    return $true
}

# Function to prompt for a valid input file path
function Get-ValidInputFilePath {
    param (
        [string]$PromptMessage,
        [string]$DefaultPath
    )
    do {
        $filePath = Read-Host "$PromptMessage (Default: $DefaultPath)"
        if ([string]::IsNullOrWhiteSpace($filePath)) {
            $filePath = $DefaultPath
        }
        if (-not (Test-Path $filePath)) {
            Write-Host "File not found: $filePath. Please enter a valid path." -ForegroundColor Red
        }
    } while (-not (Test-Path $filePath))

    # Verification step
    Write-Host "You've selected the following file: $filePath" -ForegroundColor Yellow
    $confirmation = Read-Host "Do you want to proceed with this file? (Y/N)"
    if ($confirmation -ne 'Y') {
        return Get-ValidInputFilePath $PromptMessage $DefaultPath
    }

    return $filePath
}

Clear

$MaximumFunctionCount = 32768

# Check if required modules are installed
$requiredModules = @("Microsoft.Graph.Users", "Microsoft.Graph.Groups")
foreach ($module in $requiredModules) {
    if (-not (Get-Module -ListAvailable -Name $module)) {
        Write-ErrorLog "Required module not found: $module. Please install it using 'Install-Module $module -Force'"
        exit
    }
}

Import-Module Microsoft.Graph.Users
Import-Module Microsoft.Graph.Groups

# Ask user to specify tenant type
do {
    $tenantType = Read-Host "Enter tenant type (1 for Standard/ 2 for GCC High)"
} while ($tenantType -notmatch '^(1|2)$')

# Set the appropriate environment based on tenant type
$environment = if ($tenantType -eq '2') { 'USGov' } else { 'Global' }

# Connect to Microsoft Graph
try {
    Connect-MgGraph -Scopes "Group.ReadWrite.All", "User.ReadWrite.All" -Environment $environment
}
catch {
    Write-ErrorLog "Failed to connect to Microsoft Graph" $_.Exception.Message
    exit
}

# Prompt for input file path with default value
$defaultPath = "C:\Temp\Users.csv"
$inputFilePath = Get-ValidInputFilePath "Enter the full path to the input CSV file (If nothing is entered, C:\Temp\Users.csv will be used)" $defaultPath

# Validate the CSV content structure
if (-not (Test-InputFile $inputFilePath)) {
    exit
}

# Read the CSV file containing user and group information
$users = Import-Csv -Path $inputFilePath

# Create an array to store the output data
$outputData = @()

# Initialize counters for summary
$totalUsers = 0
$successCount = 0
$alreadyMemberCount = 0
$errorCount = 0

# Main processing loop
foreach ($user in $users) {
    $totalUsers++
    $userPrincipalName = $user.userPrincipalName
    $groupName = $user.SecurityGroup

    # Section 1: Get the group ID
    try {
        $group = Get-MgGroup -Filter "displayName eq '$groupName'"
        if ($null -eq $group) {
            throw "Group not found"
        }
        $groupId = $group.Id
        Write-Host "The ID of the group '$groupName' is: $groupId"
    }
    catch {
        $errorCount++
        $message = "Error finding group '$groupName' for user '$userPrincipalName': $_"
        Write-ErrorLog $message
        $outputData += [PSCustomObject]@{
            UserPrincipalName = $userPrincipalName
            SecurityGroup = $groupName
            Status = "Error"
            Message = $message
        }
        continue
    }

    # Section 2: Get the user ID
    try {
        $userObj = Get-MgUser -Filter "userPrincipalName eq '$userPrincipalName'"
        if ($null -eq $userObj) {
            throw "User not found"
        }
        $userId = $userObj.Id
        Write-Host "The ID of the user '$userPrincipalName' is: $userId"
    }
    catch {
        $errorCount++
        $message = "Error finding user '$userPrincipalName': $_"
        Write-ErrorLog $message
        $outputData += [PSCustomObject]@{
            UserPrincipalName = $userPrincipalName
            SecurityGroup = $groupName
            Status = "Error"
            Message = $message
        }
        continue
    }

    # Section 3: Check if the user is already a member of the group
    try {
        $existingMember = Get-MgGroupMember -GroupId $groupId | Where-Object { $_.Id -eq $userId }
        if ($null -ne $existingMember) {
            $alreadyMemberCount++
            $message = "Already a member of $groupName."
            Write-Host $message -ForegroundColor Yellow
            $outputData += [PSCustomObject]@{
                UserPrincipalName = $userPrincipalName
                SecurityGroup = $groupName
                Status = "Already Member"
                Message = $message
            }
            continue
        }
    }
    catch {
        $errorCount++
        $message = "Error checking group membership for user '$userPrincipalName' in group '$groupName': $_"
        Write-ErrorLog $message
        $outputData += [PSCustomObject]@{
            UserPrincipalName = $userPrincipalName
            SecurityGroup = $groupName
            Status = "Error"
            Message = $message
        }
        continue
    }

    # Section 4: Add the user to the group
    try {
        New-MgGroupMember -GroupId $groupId -DirectoryObjectId $userId
        $successCount++
        $message = "Successfully added user '$userPrincipalName' to group '$groupName'"
        Write-Host $message -ForegroundColor Green
        $outputData += [PSCustomObject]@{
            UserPrincipalName = $userPrincipalName
            SecurityGroup = $groupName
            Status = "Success"
            Message = $message
        }
    }
    catch {
        $errorCount++
        $message = "Failed to add user '$userPrincipalName' to group '$groupName': $_"
        Write-ErrorLog $message
        $outputData += [PSCustomObject]@{
            UserPrincipalName = $userPrincipalName
            SecurityGroup = $groupName
            Status = "Error"
            Message = $_.Exception.Message 
        }
    }
}

# Section 5: Generate output file name with tenant name and timestamp
try {
    $tenantInfo = Get-MgOrganization | Select-Object -First 1
    $tenantName = if ($tenantInfo -and $tenantInfo.DisplayName) { 
        [System.Text.RegularExpressions.Regex]::Replace($tenantInfo.DisplayName, '[^\w\-\.]', '_') 
    } else { 
        'UnknownTenant' 
    }

    $timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
    $outputFilePath = "C:\temp\ResultsAddingUsersToGroups_{0}_{1}.csv" -f $tenantName, $timestamp
    $outputData | Export-Csv -Path $outputFilePath -NoTypeInformation

    Write-Host "`nOutput has been saved to: `"$outputFilePath`"" -ForegroundColor Green
}
catch {
    Write-ErrorLog "Error generating output file" $_.Exception.Message
}

# Display summary statistics on screen.
Write-Host "`nSummary:" -ForegroundColor Cyan 
Write-Host ("Total users processed: {0}" -f $totalUsers) -ForegroundColor Cyan 
Write-Host ("Successful additions: {0}" -f $successCount) -ForegroundColor Green 
Write-Host ("Already members: {0}" -f $alreadyMemberCount) -ForegroundColor Yellow 
Write-Host ("Errors: {0}" -f $errorCount) -ForegroundColor Red 

# Disconnect from Microsoft Graph  
Disconnect-MgGraph