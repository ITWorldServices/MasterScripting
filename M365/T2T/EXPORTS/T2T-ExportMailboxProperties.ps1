# Ensure you have the Exchange Online PowerShell module installed
# If not installed, run this command:
# Install-Module -Name ExchangeOnlineManagement
# Connect to Exchange Online
Connect-ExchangeOnline
# Set the path for the CSV file to save the exported mailbox properties
$exportPath = "C:\temp\fseMailboxProperties.csv"
# Get a list of all mailboxes in the organization
$mailboxes = Get-Mailbox -ResultSize Unlimited
# Create an array to store mailbox properties
$mailboxPropertiesArray = @()
# Loop through each mailbox and collect all properties
foreach ($mailbox in $mailboxes) {
    $mailboxProperties = $mailbox | Get-Mailbox | Select-Object *
    $mailboxPropertiesArray += $mailboxProperties
}
# Export mailbox properties to CSV
$mailboxPropertiesArray | Export-Csv -Path $exportPath -NoTypeInformation
# Disconnect the Exchange Online session
# Disconnect-ExchangeOnline