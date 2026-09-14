# SOURCE Tenant

# Connect to SharePoint Online
Connect-SPOService -Url "https://primocenter-admin.sharepoint.com/"


# Path to the CSV file containing user emails. The CSV must be present in the C:\Temp directory. The syntax for the users are username_domainname_com
$csvPath = "C:\temp\primocenter-OneDrive.csv"

# Import CSV data
$csvData = Import-Csv -Path $csvPath

# Loop through each user in the CSV and add migrationwiz account as a Site Collection Administrator. 
foreach ($user in $csvData) {
    $userEmail = $user.Email
    Write-Host $userEmail
    Set-SPOUser -Site https://primocenter-my.sharepoint.com/personal/$userEmail -LoginName "migrationwiz@primocenter.onmicrosoft.com" -IsSiteCollectionAdmin $true
}

Write-Host "Site Collection Administrator added for all users in the CSV."

Disconnect-SPOService



