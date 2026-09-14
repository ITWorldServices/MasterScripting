<#
.SYNOPSIS
    Export OneDrive storage usage in Microsoft 365 to CSV file with PowerShell, including provisioning status and OneDrive URL.

.DESCRIPTION
    This script connects to Microsoft Graph, retrieves all users, fetches their OneDrive usage details,
    checks if their OneDrive is provisioned, retrieves the OneDrive URL, and exports the results to a CSV file.
    It labels users with OneDrive but no license as "Unlicensed".
    It also handles tenant settings for concealed user data in reports.

.NOTES
    Requires Microsoft Graph PowerShell SDK.
    Requires "User.Read.All", "Reports.Read.All", and "Sites.Read.All" permissions.
    CSV file is saved at "C:\Temp\<TenantName>-OneDriveSizeReport-<Timestamp>.csv".
#>

$MaximumFunctionCount = 32768  # Add this line first

# Connect using required permissions
Connect-MgGraph -Scopes "User.Read.All", "Reports.Read.All", "Sites.Read.All"

# Get tenant name dynamically (first part of default domain)
$TenantName = (Get-MgDomain | Where-Object {$_.IsDefault -eq $true}).Id.Split('.')[0]

# Create timestamp for unique file name
$TimeStamp = Get-Date -Format "yyyy-MM-dd_HH-mm-ss"

# Define file paths with tenant name and timestamp
$CSVOutputFile = "C:\Temp\$TenantName-OneDriveSizeReport-$TimeStamp.csv"
$TempExportFile = "C:\Temp\TempExportFile.csv" 

# Remove the temporary export file if it exists
if (Test-Path $TempExportFile) {
    Remove-Item $TempExportFile
}

# Check if tenant reports have concealed user data, and adjust settings if necessary
$ConcealedFlag = $false

#if ((Get-MgAdminReportSetting).DisplayConcealedNames -eq $true) {
#    $Parameters = @{ displayConcealedNames = $false }
#    Write-Host "Unhiding concealed report data to retrieve full user information..." -ForegroundColor Cyan
#    Update-MgAdminReportSetting -BodyParameter $Parameters
#    $ConcealedFlag = $true
#} else {
#    Write-Host "User data is already fully visible in the reports." -ForegroundColor Cyan
#}

# Retrieve user account information
Write-Host "Fetching user account details from Microsoft Graph..." -ForegroundColor Cyan
$Properties = 'Id', 'displayName', 'userPrincipalName', 'city', 'country', 'department', 'jobTitle', 'officeLocation', 'assignedLicenses'
$UserParams = @{
    All = $true
    Filter = "userType eq 'Member'"  # Get all members, licensed or not
    ConsistencyLevel = 'Eventual'
    CountVariable = 'UserCount'
    Sort = 'displayName'
}
$Users = Get-MgUser @UserParams -Property $Properties | Select-Object -Property $Properties

# Create a hashtable to map UPNs to user details
$UserHash = @{}
foreach ($User in $Users) {
    $UserHash[$User.userPrincipalName] = $User
}

# --- MODIFIED SECTION: Check OneDrive provisioning status, URL, and license state ---
Write-Host "Checking OneDrive provisioning status, URL, and license state for each user..." -ForegroundColor Cyan
$OneDriveStatusHash = @{}
$OneDriveUrlHash = @{}
$userIndex = 0

foreach ($User in $Users) {
    $userIndex++
    $upn = $User.userPrincipalName

    # Determine if user is licensed (has any assigned licenses)
    $isLicensed = $false
    if ($User.assignedLicenses -and $User.assignedLicenses.Count -gt 0) {
        $isLicensed = $true
    }

    # Get OneDrive drive info if provisioned
    $drive = Get-MgUserDefaultDrive -UserId $User.Id -ErrorAction SilentlyContinue

    if ($drive) {
        if ($isLicensed) {
            $OneDriveStatusHash[$upn] = "OneDrive is provisioned."
            $OneDriveUrlHash[$upn] = $drive.webUrl
            Write-Host "[$userIndex/$($Users.Count)] $upn -> Provisioned (Licensed). URL: $($drive.webUrl)" -ForegroundColor Green
        } else {
            # User has OneDrive but no license
            $OneDriveStatusHash[$upn] = "Unlicensed"
            $OneDriveUrlHash[$upn] = $drive.webUrl
            Write-Host "[$userIndex/$($Users.Count)] $upn -> Provisioned (Unlicensed). URL: $($drive.webUrl)" -ForegroundColor Magenta
        }
    } else {
        if ($isLicensed) {
            $OneDriveStatusHash[$upn] = "Not Provisioned"
            $OneDriveUrlHash[$upn] = ""
            Write-Host "[$userIndex/$($Users.Count)] $upn -> Not Provisioned (Licensed)" -ForegroundColor Yellow
        } else {
            $OneDriveStatusHash[$upn] = "Unlicensed"
            $OneDriveUrlHash[$upn] = ""
            Write-Host "[$userIndex/$($Users.Count)] $upn -> No OneDrive (Unlicensed)" -ForegroundColor Magenta
        }
    }
}

# Retrieve OneDrive for Business site usage details for the last 30 days and export to a temporary CSV file
Write-Host "Retrieving OneDrive for Business site usage details..." -ForegroundColor Cyan
Get-MgReportOneDriveUsageAccountDetail -Period D30 -Outfile $TempExportFile

# Import the data from the temporary CSV file
$ODFBSites = Import-CSV $TempExportFile | Sort-Object 'User display name'
if (-not $ODFBSites) {
    Write-Host "No OneDrive sites found." -ForegroundColor Yellow
    return
}

# Calculate total storage used by all OneDrive for Business accounts
$TotalODFBGBUsed = [Math]::Round(($ODFBSites.'Storage Used (Byte)' | Measure-Object -Sum).Sum / 1GB, 2)

# Initialize a list to store report data
$Report = [System.Collections.Generic.List[Object]]::new()

# Populate the report with detailed information for each OneDrive site
foreach ($Site in $ODFBSites) {
    $UserData = $UserHash[$Site.'Owner Principal name']
    $upn = $Site.'Owner Principal name'
    $ReportLine = [PSCustomObject]@{
        Owner                   = $Site.'Owner display name'
        UserPrincipalName       = $upn
        SiteId                  = $Site.'Site Id'
        IsDeleted               = $Site.'Is Deleted'
        LastActivityDate        = $Site.'Last Activity Date'
        FileCount               = [int]$Site.'File Count'
        ActiveFileCount         = [int]$Site.'Active File Count'
        QuotaGB                 = [Math]::Round($Site.'Storage Allocated (Byte)' / 1GB, 2)
        UsedGB                  = [Math]::Round($Site.'Storage Used (Byte)' / 1GB, 2)
        PercentUsed             = if ($Site.'Storage Allocated (Byte)' -ne 0) {
                                     [Math]::Round($Site.'Storage Used (Byte)' / $Site.'Storage Allocated (Byte)' * 100, 2)
                                 } else { 0 }
        City                    = $UserData.city
        Country                 = $UserData.country
        Department              = $UserData.department
        JobTitle                = $UserData.jobTitle
        OneDriveProvisionedStatus = $OneDriveStatusHash[$upn]
        OneDriveUrl             = $OneDriveUrlHash[$upn]
    }
    $Report.Add($ReportLine)
}

# Export the report to a CSV file and display the data in a grid view
$Report | Sort-Object UsedGB -Descending | Export-CSV -NoTypeInformation -Encoding utf8 $CSVOutputFile
$Report | Sort-Object UsedGB -Descending | Out-GridView -Title "OneDrive Usage Report"

Write-Host ("Current OneDrive for Business storage consumption is {0} GB. Report saved to {1}" -f $TotalODFBGBUsed, $CSVOutputFile) -ForegroundColor Cyan

# Reset tenant report data concealment setting if it was modified earlier
if ($ConcealedFlag -eq $true) {
    Write-Host "Re-enabling data concealment in tenant reports..." -ForegroundColor Cyan
    $Parameters = @{ displayConcealedNames = $true }
    Update-MgAdminReportSetting -BodyParameter $Parameters
}

# Clean up the temporary export file
if (Test-Path $TempExportFile) {
    Remove-Item $TempExportFile
    Write-Host "Temporary export file removed." -ForegroundColor Cyan
}

Disconnect-MgGraph