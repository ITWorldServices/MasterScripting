# Connect to Exchange Online
Connect-ExchangeOnline

# Set your onmicrosoft.com domain
$onmicrosoftDomain = "PQMINC1.onmicrosoft.com"  # <-- Replace with your tenant's domain

# Add the onmicrosoft.com alias to all users
Get-Mailbox -ResultSize Unlimited | ForEach-Object {
    $alias = $_.Alias
    $newAddress = "$alias@$onmicrosoftDomain"
    Write-Host "Adding $newAddress to $($alias)..."
    Set-Mailbox $_.Identity -EmailAddresses @{add=$newAddress}
}

# Disconnect from Exchange Online
Disconnect-ExchangeOnline -Confirm:$false
