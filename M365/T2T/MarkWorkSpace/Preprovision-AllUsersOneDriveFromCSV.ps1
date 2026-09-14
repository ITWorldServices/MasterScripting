#PreProvision All Licensed Accounts 

#DESTINATION Tenant

Connect-SPOService -Url "https://primocenter-admin.sharepoint.com"
# Connect-SPOService -Url "https://ADDISONPRECISION-admin.sharepoint.us" -Region ITAR
# Connect-ExchangeOnline -UserPrincipalName migrationwiz@icsholdingllc.onmicrosoft.com

# Define the path to the CSV file. Syntax is email address
$csvPath = "C:\temp\primocenter-emailaddresses.csv"

# Import CSV data
$csvData = Import-Csv -Path $csvPath

# Loop through each row in the CSV
foreach ($row in $csvData) {
    # Access the user email from the CSV (replace "ColumnWithEmail" with the actual column name)
    $userEmail = $row.Email
    write-host $userEmail
    # Build the command to preprovision OneDrive for the user
    $command = "Request-SPOPersonalSite -UserEmails $userEmail"


    # Execute the command
    Invoke-Expression -Command $command
}

Write-Host "OneDrive preprovisioning completed for all users in the CSV."

Disconnect-SPOService
