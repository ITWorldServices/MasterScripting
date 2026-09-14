# Ensure you have the Exchange Online PowerShell module installed
# If not installed, run this command:
# Install-Module -Name ExchangeOnlineManagement
# Connect to Exchange Online
# Connect-ExchangeOnline
# Set the path for the CSV file to save the exported mailbox information
$exportPath = "C:\temp\FESMailboxStats.csv"
# Get a list of all mailboxes in the organization
$mailboxes = Get-Mailbox -ResultSize Unlimited
# Create an array to store mailbox information
$mailboxInfoArray = @()
# Loop through each mailbox and collect mailbox information
foreach ($mailbox in $mailboxes) {
    $mailboxInfo = @{
        "DisplayName" = $mailbox.DisplayName
        "UserPrincipalName" = $mailbox.UserPrincipalName
        "EmailAddress" = $mailbox.PrimarySmtpAddress
        "RecipientType" = $mailbox.RecipientTypeDetails
        "IsMailboxEnabled" = $mailbox.Enabled
        "ItemCount" = (Get-MailboxStatistics -Identity $mailbox.UserPrincipalName).ItemCount
        "TotalItemSize" = (Get-MailboxStatistics -Identity $mailbox.UserPrincipalName).TotalItemSize
        "Database" = $mailbox.Database
        "ArchiveDatabase" = $mailbox.ArchiveDatabase
        "LitigationHoldEnabled" = $mailbox.LitigationHoldEnabled
        "LitigationHoldDuration" = $mailbox.LitigationHoldDuration
        "LitigationHoldDate" = $mailbox.LitigationHoldDate
        "RetentionPolicy" = $mailbox.RetentionPolicy
        "ForwardingAddress" = $mailbox.ForwardingAddress
        "ForwardingSmtpAddress" = $mailbox.ForwardingSmtpAddress
       # "DelegateList" = ($mailbox | Get-MailboxPermission | Where-Object { $_.AccessRights -like "*FullAccess*" -and $_.IsInherited -eq $false } | Select-Object -ExpandProperty User | Get-Recipient | Select-Object -ExpandProperty UserPrincipalName)
    }
        $mailboxInfoArray += New-Object PSObject -Property $mailboxInfo
}
# Export mailbox information to CSV
$mailboxInfoArray | Export-Csv -Path $exportPath -NoTypeInformation
# Disconnect the Exchange Online session
# Disconnect-ExchangeOnline