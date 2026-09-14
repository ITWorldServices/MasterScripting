<#
.SYNOPSIS
Bulk-adds the onmicrosoft.com email alias to all Microsoft 365 user accounts listed in a CSV file.

.DESCRIPTION
This script connects to Exchange Online, imports a list of user principal names (UPNs) from a CSV file, 
dynamically retrieves your tenant's onmicrosoft.com domain, and adds an onmicrosoft.com alias to each user.
The script includes error handling for missing domains and provides confirmation messages for each alias added.

- No hardcoded tenant domain: The script automatically detects your onmicrosoft.com domain.
- Error handling: Stops if the domain cannot be found.
- Audit trail: Outputs confirmation for each alias added.
- Clean disconnection: Ensures the Exchange Online session is closed at the end.

Requirements:
- Exchange Online PowerShell V2 module
- Administrative permissions
- CSV file with a 'UserPrincipalName' column located at c:\temp
#>

# Connect to Exchange Online
Connect-ExchangeOnline
$ConfirmPreference = 'None'


# Import the CSV file
$users = Import-Csv "C:\temp\chaparral-UPNs.csv"

# Dynamically get the onmicrosoft.com domain
$onMicrosoftDomain = (Get-AcceptedDomain | Where-Object {$_.DomainName -like "*.onmicrosoft.com"} | Select-Object -First 1).DomainName

if (-not $onMicrosoftDomain) {
    Write-Error "Could not find your onmicrosoft.com domain!"
    Disconnect-ExchangeOnline
    return
}

foreach ($user in $users) {
    $upn = $user.UserPrincipalName
    # Extract the username part before the @
    $aliasPrefix = $upn.Split("@")[0]
    $alias = "$aliasPrefix@$onMicrosoftDomain"

    # Add the alias
    Set-Mailbox -Identity $upn -EmailAddresses @{add=$alias}
    Write-Host "Added alias $alias to $upn" -ForegroundColor Cyan
}

Write-host ""
Write-host "Job Completed" -ForegroundColor Green
Disconnect-ExchangeOnline