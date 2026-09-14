# ADD SPECIFIED USERS AS SITE ADMIN TO ALL ONEDRIVES LISTED ON A CSV

# SOURCE Tenant
# Connect to SharePoint Online
# Connect-SPOService -Url https://integrated3d-admin.sharepoint.com
Connect-SPOService -Url "https://eraind-admin.sharepoint.us" -Region ITAR

# Path to the CSV file containing user emails. The CSV must be present in the C:\Temp directory. The syntax for the users are username_domainname_com
# CSV header is Email and all the dots and @ symbol need to be changed to _. Example: john.smith@abc.com = john_smith_abc_com
$csvPath = "C:\temp\era-onedrive.csv"

# Import CSV data
$csvData = Import-Csv -Path $csvPath

# Loop through each user in the CSV and add the migrationwiz account as aSite Collection Administrator. 
foreach ($user in $csvData) {
    $userEmail = $user.Email
    Set-SPOUser -Site https://eraind-my.sharepoint.us/personal/$userEmail -LoginName "migrationwiz@eraind.onmicrosoft.us" -IsSiteCollectionAdmin $true
}

Write-Host "Site Collection Administrator added for all users in the CSV."

Disconnect-SPOService