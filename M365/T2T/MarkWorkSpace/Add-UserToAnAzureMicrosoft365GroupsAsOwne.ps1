# ADD OWNER to a SHAREPOINT SITE (Azure Group)
# Install Microsoft Graph module if not already installed
Install-Module Microsoft.Graph -Scope CurrentUser

# Connect with required scopes
Connect-MgGraph -Scopes "Group.ReadWrite.All", "User.Read.All"


# Get user ObjectId (if needed)
$User = Get-MgUser -UserId "migrationwiz@addisonprecision.onmicrosoft.com"

# Add user to group using group ID
New-MgGroupOwnerByRef -GroupId "3af27768-6b80-45b7-bf80-d1267ab7980f" -BodyParameter @{
    "@odata.id" = "https://graph.microsoft.com/v1.0/users/$($User.Id)"
}


# Disconnect
Disconnect-MgGraph