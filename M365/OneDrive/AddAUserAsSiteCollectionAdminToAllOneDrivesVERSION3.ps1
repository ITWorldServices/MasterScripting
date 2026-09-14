# ADD SPECIFIED USER AS SITE ADMIN TO ALL ONEDRIVES LISTED ON A CSV AND EXPORT DETAILED SUMMARY
<#
.SYNOPSIS
Adds a specified user as a site admin to multiple OneDrive sites listed in a CSV file and exports a detailed summary of the operation.

.DESCRIPTION
This PowerShell script performs the following tasks:
1. Connects to SharePoint Online (SPO) for either Standard or GCC High tenants.
2. Reads a list of user emails from a CSV file.
3. Adds a specified user as a site collection admin to each user's OneDrive.
4. Generates a detailed report of the operation, including successes, failures, and reasons.
5. Exports the results to a CSV file.

The script includes error handling for various scenarios such as:
- Invalid email formats
- OneDrive sites that are not provisioned
- Users who are already site collection admins

.PARAMETER None
This script does not accept parameters. All inputs are provided through user prompts.

.INPUTS
- Tenant type (Standard or GCC High)
- SharePoint Online admin URL
- Email of the user to be added as site collection admin
- CSV file path containing list of user emails (default: C:\temp\OneDrive.csv)

.OUTPUTS
- Detailed CSV report (default location: C:\temp\)
- Console output showing progress and summary

.NOTES
- Requires the Microsoft.Online.SharePoint.PowerShell module
- CSV file must have an 'Email' column
- Script will prompt for SharePoint Online credentials

.EXAMPLE
.\Add-SiteAdminToOneDrives.ps1

Run the script and follow the prompts to add a site admin to multiple OneDrive sites.
#>

Clear-Host

# Function to check if the module is installed and import it
function Import-RequiredModule {
    param ([string]$ModuleName)
    
    if (!(Get-Module -ListAvailable -Name $ModuleName)) {
        Write-Host "The $ModuleName module is not installed. Please install it and try again." -ForegroundColor Red
        exit
    }
    Import-Module $ModuleName
}

# Function to validate email format
function Is-ValidEmail {
    param ([string]$Email)
    return $Email -match '^[\w-\.]+@([\w-]+\.)+[\w-]{2,4}$'
}

# Function to generate timestamp
function Generate-Timestamp {
    return (Get-Date).ToString("yyyyMMdd_HHmmss")
}

# Function to export detailed summary to CSV
function Export-DetailedSummaryToCsv {
    param (
        [int]$successCount,
        [int]$alreadyAdminCount,
        [int]$notProvisionedCount,
        [int]$failureCount,
        [array]$detailedResults,
        [string]$exportPath
    )

    # Create summary data
    $summaryData = @(
        [PSCustomObject]@{ Operation = 'Successful additions'; Count = $successCount },
        [PSCustomObject]@{ Operation = 'Already site collection admin'; Count = $alreadyAdminCount },
        [PSCustomObject]@{ Operation = 'OneDrive not provisioned'; Count = $notProvisionedCount },
        [PSCustomObject]@{ Operation = 'Failed operations'; Count = $failureCount }
    )

    # Export detailed results to CSV (append to the same file)
    $detailedResults | Export-Csv -Path $exportPath -NoTypeInformation -Encoding UTF8
}

# Check and import required module
Import-RequiredModule "Microsoft.Online.SharePoint.PowerShell"

# Prompt for tenant type
do {
    $tenantType = Read-Host "Select tenant type: 1 for Standard, 2 for GCC High"
} while ($tenantType -ne "1" -and $tenantType -ne "2")
    Write-Host "Selected tenant type: " -ForegroundColor Red -NoNewline

# Display selected tenant type in blue
if ($tenantType -eq "1") {
    Write-Host "Standard" -ForegroundColor Green
    $domainSuffix = "sharepoint.com"
} else {
    Write-Host "GCC High" -ForegroundColor Yellow
    $domainSuffix = "sharepoint.us"
}

# Prompt for SPO service URL
    Write-Host "Enter the SPO service URL " -ForegroundColor Red -NoNewline
    Write-Host "(e.g., https://contoso-admin.$domainSuffix)" -ForegroundColor Yellow -NoNewline
    $spoServiceUrl = Read-Host " "
    
# Validate SPO service URL format
if ($spoServiceUrl -notmatch "^https://.*-admin\.$([regex]::Escape($domainSuffix))$") {
    Write-Host "Invalid SPO service URL format. It should be in the form 'https://contoso-admin.$domainSuffix'" -ForegroundColor Red
    exit
}

# Extract tenant name from SPO service URL and remove 'https://' and '-admin'
$tenantName = ($spoServiceUrl -replace 'https://', '') -replace "-admin\.$([regex]::Escape($domainSuffix))", ''

# Prompt for the user to be added as site collection admin
do {
    $adminToAdd = Read-Host "Enter the email of the user to be added as site collection admin"
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
$csvPath = "C:\temp\OneDrive.csv"

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

# Generate timestamp for export filename
$timestamp = Generate-Timestamp

# Define the export path with tenant name, type, and timestamp
$tenantTypeAbbr = if ($tenantType -eq "1") { "STD" } else { "GCCH" }
$exportPath = "C:\temp\AddSiteAdminToOneDrive_${tenantName}_${tenantTypeAbbr}_${timestamp}.csv"

# Initialize counters and detailed results array
$successCount = 0
$failureCount = 0
$alreadyAdminCount = 0
$notProvisionedCount = 0
$detailedResults = @()

# Loop through each user in the CSV and add the specified user as a Site Collection Administrator
foreach ($user in $csvData) {
    $userEmail = $user.Email
    $result = ""

    # Validate email format
    if ($userEmail -notmatch '^[a-zA-Z0-9._]+$') {
        $result = "Invalid email format"
        Write-Host "Invalid email format for user: $userEmail. Skipping this user." -ForegroundColor Yellow
        $failureCount++
    } else {
        $oneDriveUrl = "https://$tenantName-my.$domainSuffix/personal/$userEmail"
        
        try {
            # Check if the user's OneDrive site exists by trying to get a user object from it.
            $currentUser = Get-SPOUser -Site $oneDriveUrl -LoginName $adminToAdd -ErrorAction Stop
            
            # Check if the user is already a site collection admin.
            if ($currentUser.IsSiteAdmin) {
                $result = "Already a Site Collection Admin"
                Write-Host "$adminToAdd is already a Site Collection Admin for $userEmail" -ForegroundColor Cyan
                $alreadyAdminCount++
            } else {
                # Add the user as a site collection admin.
                Set-SPOUser -Site $oneDriveUrl -LoginName $adminToAdd -IsSiteCollectionAdmin $true -ErrorAction Stop 
                $result = "Successfully added as Site Collection Admin"
                Write-Host "Successfully added $adminToAdd as Site Collection Admin for $userEmail" -ForegroundColor Green 
                $successCount++
            }
        } catch {
            # Check if the error is due to OneDrive not being provisioned.
            if ($_.Exception.Message -like "*does not exist*") { 
                $result = "OneDrive not provisioned"
                Write-Host "OneDrive is not provisioned for user: $userEmail. Skipping this user." -ForegroundColor Yellow 
                $notProvisionedCount++
            } else { 
                $result = "Failed to process: $_" 
                Write-Host "Failed to process $userEmail. Error: $_" -ForegroundColor Red 
                $failureCount++
            }
        }
    }

   # Add the result to detailed results 
   $detailedResults += [PSCustomObject]@{
       User   = $userEmail   # User's email address.
       Result = $result      # Result of processing this user's request.
   }
}

# Display summary of operations performed.
Write-Host "`nOperation Summary:" -ForegroundColor Cyan 
Write-Host "Successful additions: $successCount" -ForegroundColor Green  
Write-Host "Already site collection admin: $alreadyAdminCount" -ForegroundColor Cyan  
Write-Host "OneDrive not provisioned: $notProvisionedCount" -ForegroundColor Yellow  
Write-Host "Failed operations: $failureCount" -ForegroundColor Red 

# Call the export function with summary counts and detailed results.
Export-DetailedSummaryToCsv `
  -successCount           $successCount `
  -alreadyAdminCount      $alreadyAdminCount `
  -notProvisionedCount     $notProvisionedCount `
  -failureCount           $failureCount `
  -detailedResults         $detailedResults `
  -exportPath             $exportPath

# Display confirmation message.
Write-Host "`nDetailed operation summary exported to: `"$exportPath`"" -ForegroundColor Green 

# Disconnect from SharePoint Online.
Disconnect-SPOService  
Write-Host "`nDisconnected from SharePoint Online." -ForegroundColor Yellow  