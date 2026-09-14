<#
.SYNOPSIS
    Bulk creates shared mailboxes in Exchange Online and assigns permissions from a CSV file.

.DESCRIPTION
    This script connects to Exchange Online, reads mailbox and permission data from a CSV file,
    creates shared mailboxes if they do not already exist, and assigns FullAccess or SendAs
    permissions as specified. At the end, it disconnects from Exchange Online to clean up the session.

    CSV columns required:
      - Name (shared mailbox display name)
      - PrimarySMTPAddress (email address of the shared mailbox)
      - User (user to assign permissions to; optional for mailbox creation)
      - AccessRights ("FullAccess" or "SendAs"; optional for mailbox creation)

.NOTES
    Author: [Your Name]
    Date:   [Today's Date]
    Tested: Exchange Online PowerShell V2 module (EXO V2)
#>

# Connect to Exchange Online
Connect-ExchangeOnline

# Import mailbox and permission data from CSV
$mailboxes = Import-Csv -Path "C:\temp\Chaparral-SharedMailboxes.csv"

foreach ($mbx in $mailboxes) {
    $mailboxName = $mbx.Name.Trim()
    
    # Check if mailbox already exists
    $existing = Get-Mailbox -Filter "PrimarySmtpAddress -eq '$($mbx.PrimarySMTPAddress)'" -ErrorAction SilentlyContinue
    
    if (-not $existing) {
        try {
            New-Mailbox -Name $mailboxName `
                        -PrimarySmtpAddress $mbx.PrimarySMTPAddress `
                        -Shared `
                        -ErrorAction Stop
            Write-Host "Created shared mailbox: $mailboxName" -ForegroundColor Green
        }
        catch {
            Write-Host "Error creating ${mailboxName}: $_" -ForegroundColor Red
        }
    }
    else {
        Write-Host "Mailbox '$mailboxName' already exists." -ForegroundColor Yellow
    }

    # Assign permissions if User and AccessRights are specified
    if ($mbx.User -and $mbx.AccessRights) {
        if ($mbx.AccessRights -eq "FullAccess") {
            Add-MailboxPermission -Identity $mbx.PrimarySMTPAddress `
                                  -User $mbx.User `
                                  -AccessRights FullAccess `
                                  -Confirm:$false
            Write-Host "Granted FullAccess to $($mbx.User) on $mailboxName" -ForegroundColor Cyan
        }
        elseif ($mbx.AccessRights -eq "SendAs") {
            Add-RecipientPermission -Identity $mbx.PrimarySMTPAddress `
                                    -Trustee $mbx.User `
                                    -AccessRights SendAs `
                                    -Confirm:$false
            Write-Host "Granted SendAs to $($mbx.User) on $mailboxName" -ForegroundColor Cyan
        }
    }
}

# Disconnect from Exchange Online
Disconnect-ExchangeOnline -Confirm:$false

<#
.SUMMARY
    This script automates the creation of shared mailboxes and assignment of permissions in Exchange Online.
    It reads all required information from a CSV file, ensuring repeatable, error-minimized mailbox management.
    Remember to verify your CSV structure and test with a small dataset before running in production.
#>
