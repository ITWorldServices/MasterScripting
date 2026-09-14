# Load Exchange module
Add-PSSnapin Microsoft.Exchange.Management.PowerShell.SnapIn

# Connect to Exchange server
$ExchangeServer = "mail-al2.brifutelectric.com"
$Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri http://$ExchangeServer/PowerShell/ -Authentication Kerberos
Import-PSSession $Session

# Export distribution groups and owners
$Groups = Get-DistributionGroup -ResultSize Unlimited
$ExportData = @()

foreach ($Group in $Groups) {
    $GroupInfo = Get-DistributionGroup $Group.Name | Select-Object Name, PrimarySmtpAddress, ManagedBy
    $ExportData += $GroupInfo
}

# Export data to CSV file
$ExportData | Export-Csv -Path "C:\temp\DistributionGroupsOWNERS.csv" -NoTypeInformation

# Disconnect from Exchange server
Remove-PSSession $Session
