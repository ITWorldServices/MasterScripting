<#
.SYNOPSIS
This PowerShell script exports user information from Azure Active Directory (Azure AD), also known as Azure Entra ID, using the Microsoft Graph API. 
It provides options to export either basic or detailed user information into CSV files.

.DESCRIPTION
The script facilitates the following functionalities:
1. Connects to Microsoft Graph with the required permissions (`User.Read.All`).
2. Retrieves and formats the Azure AD tenant's display name for use in file naming.
3. Generates a timestamp to include in file names for better organization.
4. Exports user information:
   - Basic User Information: Includes essential properties such as `DisplayName`, `GivenName`, `Surname`, `UserPrincipalName`, and `Mail`.
   - Detailed User Information: Includes comprehensive properties such as job title, department, manager details, office location, phone numbers, and more.
5. Saves exported data into CSV files in a specified directory (`C:\Temp`) with filenames containing the tenant name and timestamp.
6. Provides an interactive prompt for the user to select an export type (basic, detailed, or both).
7. Optionally displays exported data in an interactive grid view window for quick review.
8. Ensures proper cleanup by disconnecting from Microsoft Graph after execution.

.EXAMPLE
Run the script with administrative privileges:
C:\Scripts\GetAzureADUSERSandEXPORT.ps1

.NOTES
- The script requires the Microsoft Graph PowerShell module.
- Ensure that you have appropriate permissions (`User.ReadWrite.All` and `Directory.ReadWrite.All`) before running the script.
- This version is intended for standard Microsoft 365 tenants, not GCC High or other specialized environments.
#>

# Clear the screen
Clear-Host

# Function to connect to Microsoft Graph
function Connect-MgGraphConnection {
    param (
        [string]$Scope
    )
    Connect-MgGraph -Scopes $Scope
}

# Function to get tenant name
function Get-TenantName {
    # Retrieve tenant details using Microsoft Graph API
    $tenantDetails = Get-MgOrganization
    return $tenantDetails.DisplayName -replace '[^a-zA-Z0-9]', '_'  # Replace invalid characters with underscores
}

# Function to generate a timestamp
function Get-Timestamp {
    return (Get-Date -Format "yyyyMMdd_HHmmss")
}

# Function to export basic user information
function Export-BasicUsers {
    param (
        [string]$CsvPath
    )
    
    # Retrieve users using Microsoft Graph API with basic properties
    $propertyParams = @{
        All = $true
        Property = @(
            'DisplayName'
            'GivenName'
            'Surname'
            'UserPrincipalName'
            'Mail'
        )
    }

    $users = Get-MgUser @propertyParams

    # Create a List to store the report
    $Report = [System.Collections.Generic.List[Object]]::new()

    # Collect and loop through all users
    foreach ($user in $users) {
        # Create a custom object for each item and add it to the report
        $ReportLine = [PSCustomObject]@{
            DisplayName       = $user.DisplayName
            GivenName         = $user.GivenName
            Surname           = $user.Surname
            UserPrincipalName = $user.UserPrincipalName
            Mail              = $user.Mail
        }

        # Add the report line to the List
        $Report.Add($ReportLine)
    }

    # Export users to CSV with proper column headers
    $Report | Sort-Object DisplayName | Export-Csv -Path $CsvPath -NoTypeInformation -Encoding UTF8

    # Display data in Out-GridView (optional)
    $Report | Out-GridView -Title "Basic Users Export"

    Write-Host "CSV file exported to: $CsvPath" -ForegroundColor Green
}

# Function to export detailed user information
function Export-DetailedUsers {
    param (
        [string]$CsvPath
    )
    
    # Retrieve users using Microsoft Graph API with detailed properties
    $propertyParams = @{
        All = $true
        ExpandProperty = 'manager'
        Property = @(
            'GivenName'
            'Surname'
            'DisplayName'
            'UserPrincipalName'
            'JobTitle'
            'Mail'
            'Department'
            'CompanyName'
            'OfficeLocation'
            'EmployeeID'
            'MobilePhone'
            'BusinessPhones'
            'StreetAddress'
            'City'
            'PostalCode'
            'State'
            'Country'
            'UserType'
            'OnPremisesSyncEnabled'
            'AccountEnabled'
            'CreatedDateTime'
        )
    }

    $users = Get-MgUser @propertyParams

    # Create a List to store the report
    $Report = [System.Collections.Generic.List[Object]]::new()

    # Collect and loop through all users
    foreach ($user in $users) {
        # Get manager information (if available)
        if ($user.Manager) {
            $managerDN = $user.Manager.AdditionalProperties.DisplayName
            $managerUPN = $user.Manager.AdditionalProperties.UserPrincipalName
        } else {
            $managerDN, $managerUPN = '', ''
        }

        # Create a custom object for each item and add it to the report
        $ReportLine = [PSCustomObject]@{
            ID                     = $user.Id
            GivenName              = $user.GivenName
            Surname                = $user.Surname
            DisplayName            = $user.DisplayName
            UserPrincipalName      = $user.UserPrincipalName
            Mail                   = $user.Mail
            JobTitle               = $user.JobTitle
            ManagerDisplayName     = $managerDN
            ManagerUserPrincipal   = $managerUPN
            Department             = $user.Department
            Company                = $user.CompanyName
            OfficeLocation         = $user.OfficeLocation
            EmployeeID             = $user.EmployeeID
            MobilePhoneNumber      = $user.MobilePhone
            BusinessPhones         = ($user.BusinessPhones -join ',')
            StreetAddress          = $user.StreetAddress
            City                   = $user.City
            PostalCode             = $user.PostalCode
            State                  = $user.State
            Country                = $user.Country
            UserType               = $user.UserType
            OnPremisesSyncEnabled  = if ($user.OnPremisesSyncEnabled) { "Enabled" } else { "Disabled" }
            AccountEnabled         = if ($user.AccountEnabled) { "Enabled" } else { "Disabled" }
            CreatedDateTime        = $user.CreatedDateTime.ToString("yyyy-MM-dd")
        }

        # Add the report line to the List
        $Report.Add($ReportLine)
    }

    # Export users to CSV with proper column headers
    $Report | Sort-Object DisplayName | Export-Csv -Path $CsvPath -NoTypeInformation -Encoding UTF8

    # Display data in Out-GridView (optional)
    $Report | Out-GridView -Title "Detailed Users Export"

    Write-Host "CSV file exported to: $CsvPath" -ForegroundColor Green
}

# Main script logic starts here

Write-Host "Azure Entra ID User Export Script"
Write-Host "-----------------------------------"

# Install Microsoft Graph PowerShell module if not installed already.
if (-not (Get-Module -ListAvailable Microsoft.Graph)) {
    Install-Module Microsoft.Graph -Force -AllowClobber 
}

# Connect to Microsoft Graph with required scope.
Connect-MgGraphConnection -Scope "User.Read.All"

# Retrieve tenant name for file naming.
$tenantName = Get-TenantName

# Prompt user for export type.
Write-Host "Select export type:"
Write-Host "1. Basic User Information"
Write-Host "2. Detailed User Information"
Write-Host "3. Both Basic and Detailed"

$choice = Read-Host "Enter your choice (1, 2, or 3)"

# Generate file paths for CSVs with tenant name.
$timestamp       = Get-Timestamp 
$CsvfileBasic   = "C:\Temp\$tenantName-BasicUsers_$timestamp.csv"
$CsvfileDetailed= "C:\Temp\$tenantName-DetailedUsers_$timestamp.csv"

switch ($choice) {
    1 { Export-BasicUsers -CsvPath $CsvfileBasic }
    2 { Export-DetailedUsers -CsvPath $CsvfileDetailed }
    3 {
        Export-BasicUsers -CsvPath $CsvfileBasic 
        Export-DetailedUsers -CsvPath $CsvfileDetailed 
    }
    Default { Write-Host "Invalid choice. Exiting." }
}

Disconnect-MgGraph

