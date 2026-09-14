# Install the Microsoft Graph module if not already installed
Install-Module Microsoft.Graph -Scope CurrentUser

# Connect to Graph
Connect-MgGraph -Scopes "User.Read.All"

# Get all users and their ImmutableId
Get-MgUser -All | Select-Object DisplayName, UserPrincipalName, ImmutableId | Format-Table
