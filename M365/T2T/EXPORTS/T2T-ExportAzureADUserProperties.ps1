# Import the AzureAD module
# Import-Module AzureAD
# Connect to Azure AD
# Connect-AzureAD
# Set the path for the CSV file to save the exported user properties
$exportPath = "C:\temp\FESAzureADUserProperties.csv"
# Get all Azure AD users
$users = Get-AzureADUser -All $true
# Create an array to store user properties
$userPropertiesArray = @()
# Loop through each user and collect all properties
foreach ($user in $users) {
    $userProperties = $user | Select-Object *
    $userPropertiesArray += $userProperties
}
# Export user properties to CSV
$userPropertiesArray | Export-Csv -Path $exportPath -NoTypeInformation
# Disconnect from Azure AD
# Disconnect-AzureAD