# Install and import Microsoft Graph module if not already done
# Install-Module Microsoft.Graph -Scope CurrentUser -Force

$MaximumFunctionCount = 32768
Import-Module Microsoft.Graph
# Import-Module Microsoft.Graph.Users
# Import-Module Microsoft.Graph.Authentication



# Connect to Microsoft Graph with necessary permissions
# Non GCC High tenant
Connect-MgGraph -Scopes "User.Read.All"

# GCC High Tenant
# Connect-MgGraph -Environment USGovDoD -Scopes "User.Read.All"

# Get all users with all properties
$users = Get-MgUser -All -Property *

# Export users to CSV file
$users | Select-Object * | Export-Csv -Path "C:\Temp\AzureADUsers_AllProperties.csv" -NoTypeInformation -Encoding UTF8

# Disconnect from Microsoft Graph
# Disconnect-MgGraph