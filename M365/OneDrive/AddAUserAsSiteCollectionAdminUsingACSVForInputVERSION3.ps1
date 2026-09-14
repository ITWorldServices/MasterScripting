<#
.SYNOPSIS
Adds a specified user as a Site Collection Administrator to multiple OneDrive sites listed in a CSV file and exports a summary of operations.

.DESCRIPTION
This script automates the process of adding a user as a Site Collection Administrator to multiple OneDrive sites. It performs the following key functions:

1. User Input: Prompts for the tenant type (standard or GCC High), SharePoint Online (SPO) service URL, and the email of the user to be added as a Site Collection Administrator. It remembers previous inputs for convenience.

2. CSV Processing: Reads a list of user emails from a CSV file located at C:\temp\Maks-OneDrive.csv.

3. OneDrive Site Administration: For each user in the CSV:
   - Constructs the OneDrive URL based on the tenant type
   - Checks if the specified admin is already a Site Collection Administrator
   - Adds the admin if they're not already one
   - Handles errors, including cases where OneDrive is not provisioned

4. Reporting: 
   - Provides real-time console output of operations
   - Generates a summary of successful additions, already existing admins, non-provisioned OneDrives, and failures
   - Exports a detailed summary to a CSV file with tenant name and timestamp

5. Error Handling: Includes comprehensive error checking and reporting throughout the process

6. SharePoint Online Integration: Connects to and disconnects from SharePoint Online service

This script is particularly useful for IT administrators managing large numbers of OneDrive sites, providing an efficient way to grant administrative access across multiple user accounts.

.PARAMETER None
This script does not accept any parameters.

.INPUTS
None. You cannot pipe objects to this script.

.OUTPUTS
This script outputs a CSV file with a summary of operations performed.

.EXAMPLE
.\Add-SiteAdminToOneDrives.ps1

.NOTES
File Name      : Add-SiteAdminToOneDrives.ps1
Prerequisite   : SharePoint Online Management Shell, appropriate permissions
#>

# ADD SPECIFIED USERS AS SITE ADMIN TO ALL ONEDRIVES LISTED ON A CSV AND EXPORT SUMMARY

Clear-Host

# Function to validate email format
function Is-ValidEmail {
    param ([string]$Email)
    return $Email -match '^[\w-\.]+@([\w-]+\.)+[\w-]{2,4}$'
}

# Function to export operation summary to CSV
function Export-SummaryToCsv {
    param (
        [int]$successCount,
        [int]$alreadyAdminCount,
        [int]$notProvisionedCount,
        [int]$failureCount,
        [string]$exportPath
    )

    $summaryData = @(
        @{ Operation = 'Successful additions'; Count = $successCount },
        @{ Operation = 'Already site collection admin'; Count = $alreadyAdminCount },
        @{ Operation = 'OneDrive not provisioned'; Count = $notProvisionedCount },
        @{ Operation = 'Failed operations'; Count = $failureCount }
    )

    $summaryData | Export-Csv -Path $exportPath -NoTypeInformation -Encoding UTF8
}

# Function to get input with previous value
function Get-InputWithPrevious {
    param (
        [string]$prompt,
        [string]$envVarName
    )

    $prevValue = [Environment]::GetEnvironmentVariable($envVarName, "User")
    if ($prevValue) {
        Write-Host "$prompt" -ForegroundColor White -NoNewline
        Write-Host " (Press Enter to use previous value: $prevValue)" -ForegroundColor Yellow
        $input = Read-Host
        if ([string]::IsNullOrWhiteSpace($input)) {
            return $prevValue
        }
    } else {
        Write-Host $prompt -ForegroundColor White -NoNewline
        $input = Read-Host
    }
    [Environment]::SetEnvironmentVariable($envVarName, $input, "User")
    return $input
}

# Prompt for tenant type
do {
    Write-Host "Choose tenant type:" -ForegroundColor White
    Write-Host "1. Standard tenant" -ForegroundColor Cyan
    Write-Host "2. GCC High tenant" -ForegroundColor Cyan
    $tenantType = Read-Host "Enter your choice (1 or 2)"
} while ($tenantType -notin @('1', '2'))

# Prompt for SPO service URL
Write-Host "Enter the SPO service URL (e.g., https://contoso-admin.sharepoint.com)" -ForegroundColor White -NoNewline
$spoServiceUrl = Get-InputWithPrevious "" "SPOServiceUrl"

# Validate SPO service URL format
if ($tenantType -eq '1' -and $spoServiceUrl -notmatch '^https://.*-admin\.sharepoint\.com$') {
    Write-Host "Invalid SPO service URL format for standard tenant. It should be in the form 'https://contoso-admin.sharepoint.com'" -ForegroundColor Red
    exit
} elseif ($tenantType -eq '2' -and $spoServiceUrl -notmatch '^https://.*-admin\.sharepoint\.us$') {
    Write-Host "Invalid SPO service URL format for GCC High tenant. It should be in the form 'https://contoso-admin.sharepoint.us'" -ForegroundColor Red
    exit
}

# Prompt for the user to be added as site collection admin
do {
    Write-Host "Enter the email of the user to be added as site collection admin" -ForegroundColor White -NoNewline
    $adminToAdd = Get-InputWithPrevious "" "AdminToAdd"
    if (-not (Is-ValidEmail $adminToAdd)) {
        Write-Host "Invalid email format. Please enter a valid email address." -ForegroundColor Yellow
    }
} while (-not (Is-ValidEmail $adminToAdd))

# Connect to SharePoint Online
try {
    Connect-SPOService -Url $spoServiceUrl -ErrorAction Stop
    Write-Host "Successfully connected to SharePoint Online." -ForegroundColor Green
} catch {
    Write-Host "Failed to connect to SharePoint Online. Error: $_" -ForegroundColor Red
    exit
}

# Path to the CSV file containing user emails
$csvPath = "C:\temp\Maks-OneDrive.csv"

# Check if the CSV file exists
if (-not (Test-Path $csvPath)) {
    Write-Host "CSV file not found at $csvPath. Please ensure the file exists." -ForegroundColor Red
    Disconnect-SPOService
    exit
}

# Import CSV data
try {
    $csvData = Import-Csv -Path $csvPath -ErrorAction Stop
} catch {
    Write-Host "Failed to import CSV file. Error: $_" -ForegroundColor Red
    Disconnect-SPOService
    exit
}

# Check if CSV is empty
if ($csvData.Count -eq 0) {
    Write-Host "The CSV file is empty. Please add user data and try again." -ForegroundColor Red
    Disconnect-SPOService
    exit
}

# Validate CSV structure
if (-not ($csvData[0].PSObject.Properties.Name -contains "Email")) {
    Write-Host "The CSV file does not contain an 'Email' column. Please check the CSV format." -ForegroundColor Red
    Disconnect-SPOService
    exit
}

# Extract the domain from the SPO service URL for constructing OneDrive URLs
$domain = ($spoServiceUrl -replace 'https://', '' -split '\.')[0] -replace '-admin', ''

# Initialize counters
$successCount = 0
$failureCount = 0
$alreadyAdminCount = 0
$notProvisionedCount = 0

# Loop through each user in the CSV and add the specified user as a Site Collection Administrator
foreach ($user in $csvData) {
    $userEmail = $user.Email

    # Validate email format
    if ($userEmail -notmatch '^[a-zA-Z0-9._]+$') {
        Write-Host "Invalid email format for user: $userEmail. Skipping this user." -ForegroundColor Yellow
        $failureCount++
        continue
    }

    # Construct OneDrive URL based on tenant type
    if ($tenantType -eq '1') {
        $oneDriveUrl = "https://$domain-my.sharepoint.com/personal/$userEmail"
    } else {
        $oneDriveUrl = "https://$domain-my.sharepoint.us/personal/$userEmail"
    }
    
    try {
        # Check if the user's OneDrive site exists by trying to get a user object from it.
        $currentUser = Get-SPOUser -Site $oneDriveUrl -LoginName $adminToAdd -ErrorAction Stop
        
        # Check if the user is already a site collection admin.
        if ($currentUser.IsSiteAdmin) {
            Write-Host "$adminToAdd is already a Site Collection Admin for $userEmail" -ForegroundColor Cyan
            $alreadyAdminCount++
            continue
        }

        # Add the user as a site collection admin.
        Set-SPOUser -Site $oneDriveUrl -LoginName $adminToAdd -IsSiteCollectionAdmin $true -ErrorAction Stop
        
        Write-Host "Successfully added $adminToAdd as Site Collection Admin for $userEmail" -ForegroundColor Green
        $successCount++
        
    } catch {
        # Check if the error is due to OneDrive not being provisioned.
        if ($_.Exception.Message -like "*does not exist*") {
            Write-Host "OneDrive is not provisioned for user: $userEmail. Skipping this user." -ForegroundColor Yellow
            $notProvisionedCount++
        } else {
            Write-Host "Failed to process $userEmail. Error: $_" -ForegroundColor Red 
            $failureCount++
        }
    }
}

# Display summary of operations performed.
Write-Host "`nOperation Summary:" -ForegroundColor Cyan
Write-Host "Successful additions: $successCount" -ForegroundColor Green 
Write-Host "Already site collection admin: $alreadyAdminCount" -ForegroundColor Cyan 
Write-Host "OneDrive not provisioned: $notProvisionedCount" -ForegroundColor Yellow 
Write-Host "Failed operations: $failureCount" -ForegroundColor Red 

# Extract tenant name from SPO service URL
$tenantName = $domain

# Create timestamp
$timestamp = Get-Date -Format "yyyyMMdd-HHmmss"

# Define the export path with tenant name and timestamp
$exportPath = "C:\temp\AddSiteAdminToOneDrive-$tenantName-$timestamp.csv"

# Call the export function with summary counts
Export-SummaryToCsv -successCount $successCount -alreadyAdminCount $alreadyAdminCount -notProvisionedCount $notProvisionedCount -failureCount $failureCount -exportPath $exportPath

# Display confirmation message
Write-Host "Operation summary exported to $exportPath" -ForegroundColor Green

# Disconnect from SharePoint Online.
Disconnect-SPOService 
Write-Host "Disconnected from SharePoint Online." -ForegroundColor Yellow