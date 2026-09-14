# ADD a OWNER and a MEMBER to ALL SHAREPOINT SITES (Azure Group)

# Connect to Microsoft Graph with required permissions
Connect-MgGraph -Scopes "Group.ReadWrite.All", "User.Read.All"

# Specify the user's UPN (e.g., "user@domain.com")
$UserUPN = "migrationwiz@addisonprecision.onmicrosoft.com"

# Get the user's Object ID
$User = Get-MgUser -Filter "userPrincipalName eq '$UserUPN'"
if (-not $User) {
    Write-Host "User $UserUPN not found." -ForegroundColor Red
    exit
}

# Get all Microsoft 365 Groups (Unified groups)
$AllGroups = Get-MgGroup -All -Filter "groupTypes/any(c:c eq 'Unified')"

# Iterate through each group
foreach ($Group in $AllGroups) {
    $GroupId = $Group.Id
    $GroupName = $Group.DisplayName

    # Check if user is already a member
    $IsMember = Get-MgGroupMember -GroupId $GroupId -Filter "id eq '$($User.Id)'" -ErrorAction SilentlyContinue
    if (-not $IsMember) {
        # Add user as a member
        New-MgGroupMember -GroupId $GroupId -DirectoryObjectId $User.Id
        Write-Host "Added $UserUPN as MEMBER to $GroupName" -ForegroundColor Green
    }
    else {
        Write-Host "$UserUPN is already a MEMBER of $GroupName" -ForegroundColor Yellow
    }

    # Check if user is already an owner
    $IsOwner = Get-MgGroupOwner -GroupId $GroupId -Filter "id eq '$($User.Id)'" -ErrorAction SilentlyContinue
    if (-not $IsOwner) {
        # Add user as an owner
        New-MgGroupOwnerByRef -GroupId $GroupId -BodyParameter @{
            "@odata.id" = "https://graph.microsoft.com/v1.0/users/$($User.Id)"
        }
        Write-Host "Added $UserUPN as OWNER to $GroupName" -ForegroundColor Cyan
    }
    else {
        Write-Host "$UserUPN is already an OWNER of $GroupName" -ForegroundColor Yellow
    }
}

# Disconnect
Disconnect-MgGraph
