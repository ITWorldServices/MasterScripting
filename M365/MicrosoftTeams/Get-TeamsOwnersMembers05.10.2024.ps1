# Install the Microsoft Teams PowerShell module if not already installed
if (-not (Get-Module -Name MicrosoftTeams -ListAvailable)) {
    Install-Module -Name MicrosoftTeams -Force -AllowClobber
}

# Connect to your Office 365 account
Connect-MicrosoftTeams

# Initialize an empty array to store all team members and owners
$TeamMembersAndOwners = @()

# Get all teams in the organization
$Teams = Get-Team

# Iterate through each team and retrieve members and owners
foreach ($Team in $Teams) {
    $TeamDisplayName = $Team.DisplayName
    $Members = Get-TeamUser -GroupId $Team.GroupId -Role Member
    $Owners = Get-TeamUser -GroupId $Team.GroupId -Role Owner

    # Iterate through members and add them to the array
    foreach ($Member in $Members) {
        $TeamMembersAndOwners += [PSCustomObject]@{
            "Team" = $TeamDisplayName
            "Role" = "Member"
            "User" = $Member.User
        }
    }

    # Iterate through owners and add them to the array
    foreach ($Owner in $Owners) {
        $TeamMembersAndOwners += [PSCustomObject]@{
            "Team" = $TeamDisplayName
            "Role" = "Owner"
            "User" = $Owner.User
        }
    }
}

# Export the combined data to a CSV file
$TeamMembersAndOwners | Export-Csv -Path "CLIENTNAME_Team_Members_And_Owners.csv" -NoTypeInformation -Force
