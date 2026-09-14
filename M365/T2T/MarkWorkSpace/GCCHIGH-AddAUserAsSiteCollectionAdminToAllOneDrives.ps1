# SOURCE Tenant

# Connect to GCC High SharePoint Online
Connect-SPOService -Url "https://chandlerindustries-admin.sharepoint.us" -Region ITAR

# Path to the CSV file containing user emails. The CSV must be present in the C:\Temp directory. The syntax for the users are username_domainname_com
$csvPath = "C:\temp\Chandler-emailaddresses-SiteCollectionAdmin.csv"

# Import CSV data
$csvData = Import-Csv -Path $csvPath

# Loop through each user in the CSV and add migrationwiz account as a Site Collection Administrator. 
foreach ($user in $csvData) {
    $userEmail = $user.Email
    Set-SPOUser -Site https://chandlerindustries-my.sharepoint.us/personal/$userEmail -LoginName "migrationwiz@chandlerindustries.onmicrosoft.us" -IsSiteCollectionAdmin $true
}

Write-Host "Site Collection Administrator added for all users in the CSV."

Disconnect-SPOService