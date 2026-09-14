
# Create an excel spreadsheet with the user email addresses listed in it. Turn it into a Table in Excel before saving it with the title Email in the email column. 
# Change the name below to match where your file is and what it name is in the $excelFilePath variable
# Change the Admin site below to the correct client admin portal site name.

# Import the required module to handle Excel files
 Import-Module ImportExcel

# Specify the path to your Excel file
$excelFilePath = "C:\temp\COMPANYNAMEOneDrive.xlsx"

# Read the Excel file and extract the 'Email' column
$emails = Import-Excel -Path $excelFilePath | Select-Object -ExpandProperty Email

# Admin site
$adminSiteURL = "https://<your-tenant>-admin.sharepoint.com"  # Replace <your-tenant> with your tenant name

# Connect to SharePoint Online
Connect-SPOService -Url $adminSiteURL


# Loop through each email in the list and request personal site
foreach ($email in $emails) {
    Write-Host "Requesting personal site for $email"
    Request-SPOPersonalSite -UserEmails $email
}

# Disconnect from the SharePoint service
Disconnect-SPOService
