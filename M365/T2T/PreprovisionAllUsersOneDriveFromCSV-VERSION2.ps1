################################################################################################################################################
# This Script will Preprovision all the users contained on a CSV in a tenant. There are options for Connecting to standard or GCC High tenants.# 
################################################################################################################################################

# Clear the screen
Clear-Host

# Function to get user input for tenant type
function Get-TenantType {
    $validInput = $false
    while (-not $validInput) {
        # Display the prompt in orange (DarkYellow)
        Write-Host "Enter tenant type (1 for Standard, 2 for GCC High):" -ForegroundColor Red
        
        # Capture user input
        $tenantType = Read-Host
        
        switch ($tenantType) {
            "1" { 
                $validInput = $true
                return "Standard"
            }
            "2" { 
                $validInput = $true
                return "GCC High"
            }
            default {
                Write-Host "Invalid input. Please enter 1 for Standard or 2 for GCC High." -ForegroundColor Yellow
            }
        }
    }
}


# Get tenant type from user
$tenantType = Get-TenantType

# Connect to the appropriate tenant based on user input
if ($tenantType -eq "Standard") {
    $tenantUrl = "https://eraind-admin.sharepoint.com"
    Connect-SPOService -Url $tenantUrl
    Write-Host "Connected to Standard tenant" -ForegroundColor Cyan
}
else {
    $tenantUrl = "https://eraind-admin.sharepoint.us"
    Connect-SPOService -Url $tenantUrl -Region ITAR
    Write-Host "Connected to GCC High tenant" -ForegroundColor Cyan
}

# Define the path to the CSV file. Syntax is email address
# The required header is "Mail"
# The email format should be: username@abc.com = username_abc_com
$csvPath = "C:\temp\era-OneDrive.csv"

# Import CSV data
$csvData = Import-Csv -Path $csvPath

# Arrays to store results
$results = @()

# Loop through each row in the CSV
foreach ($row in $csvData) {
    # Access the user email from the CSV (replace "ColumnWithEmail" with the actual column name)
    $userEmail = $row.Mail

    Write-Host "Processing: $userEmail" -ForegroundColor Cyan

    # Build the command to preprovision OneDrive for the user
    $command = "Request-SPOPersonalSite -UserEmails $userEmail -ErrorAction SilentlyContinue -ErrorVariable provisionError"

    # Execute the command
    $result = Invoke-Expression -Command $command

    # Check if there was an error
    if ($provisionError) {
        $errorMessage = $provisionError[0].Exception.Message
        $status = "Failed"
        Write-Host "Failed: $userEmail - $errorMessage" -ForegroundColor Red
    } else {
        $status = "Success"
        $errorMessage = ""
        Write-Host "Success: $userEmail" -ForegroundColor Green
    }

    # Add result to the results array
    $results += [PSCustomObject]@{
        Email = $userEmail
        Status = $status
        Error = $errorMessage
    }
}

# Output summary results
$successfulUsers = $results | Where-Object { $_.Status -eq "Success" }
$failedUsers = $results | Where-Object { $_.Status -eq "Failed" }

Write-Host "`nOneDrive preprovisioning completed." -ForegroundColor Yellow
Write-Host "Successful preprovisioning: $($successfulUsers.Count) users" -ForegroundColor Green
Write-Host "Failed preprovisioning: $($failedUsers.Count) users" -ForegroundColor Red

# Export all results to CSV
$exportPath = "C:\temp\OneDriveProvisioningResults.csv"
$results | Export-Csv -Path $exportPath -NoTypeInformation
Write-Host "All results exported to $exportPath" -ForegroundColor Yellow

# Export failed users to CSV
if ($failedUsers.Count -gt 0) {
    $failedExportPath = "C:\temp\FailedOneDriveProvisioning.csv"
    $failedUsers | Export-Csv -Path $failedExportPath -NoTypeInformation
    Write-Host "Failed users exported to $failedExportPath" -ForegroundColor Yellow
}

Disconnect-SPOService