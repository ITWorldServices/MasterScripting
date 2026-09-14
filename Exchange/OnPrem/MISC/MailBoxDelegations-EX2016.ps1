# Connect to the Exchange server
$ExchangeServer = "mail-al2.brifutelectric.com" # Replace with your Exchange server name
$Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri http://$ExchangeServer/PowerShell/ -Authentication Kerberos
Import-PSSession $Session -AllowClobber -DisableNameChecking

# Define the output file path
$OutputFile = "C:\temp\MailboxDelegations.csv" # Replace with the desired output file path

# Get all mailboxes
$Mailboxes = Get-Mailbox -ResultSize Unlimited

# Create an array to store the mailbox delegations
$MailboxDelegations = @()

# Iterate through each mailbox and retrieve the delegations
foreach ($Mailbox in $Mailboxes) {
    $MailboxIdentity = $Mailbox.Identity

    # Get the mailbox delegations
    $Delegations = Get-MailboxPermission -Identity $MailboxIdentity |
                  Where-Object { $_.AccessRights -like "*FullAccess*" -and $_.IsInherited -eq $false } |
                  Select-Object @{Name = "Mailbox"; Expression = { $MailboxIdentity }},
                                User,
                                AccessRights

    # Add the delegations to the array
    $MailboxDelegations += $Delegations
}

# Export the mailbox delegations to a CSV file
$MailboxDelegations | Export-Csv -Path $OutputFile -NoTypeInformation

# Disconnect from the Exchange server
Remove-PSSession $Session
