# Import AzureAD module
Import-Module AzureAD

# Connect specifically to the US Government cloud (GCC High)
Connect-AzureAD -AzureEnvironmentName AzureUSGovernment

# Export user DisplayName, UserPrincipalName, and ImmutableId to CSV
Get-AzureADUser -All $true | Select-Object DisplayName, UserPrincipalName, ImmutableId | Export-Csv -Path "C:\temp\ImmutableIDs.csv" -NoTypeInformation
