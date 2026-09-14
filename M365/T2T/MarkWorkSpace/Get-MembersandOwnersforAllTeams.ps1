# Connect to Microsoft Teams
Connect-MicrosoftTeams

# Prepare result array
$Result = @()

# Get all teams
$AllTeams = Get-Team
$TotalTeams = $AllTeams.Count
$i = 0

# Iterate through each team and get users
foreach ($Team in $AllTeams) {
    $i++
    Write-Progress -Activity "Fetching users for $($Team.DisplayName)" -Status "$i out of $TotalTeams completed"
    try {
        # Get users for the team
        $TeamUsers = Get-TeamUser -GroupId $Team.GroupId

        # Add each user to the result array
        foreach ($TeamUser in $TeamUsers) {
            $Result += [PSCustomObject]@{
                TeamName        = $Team.DisplayName
                TeamVisibility  = $Team.Visibility
                UserName        = $TeamUser.Name
                UserPrincipalName = $TeamUser.User
                Role            = $TeamUser.Role
                GroupId         = $Team.GroupId
            }
        }
    }
    catch {
        Write-Host "Error occurred for $($Team.DisplayName)" -ForegroundColor Yellow
        Write-Host $_ -ForegroundColor Red
    }
}

# Export the result to a well-formatted CSV file
$Result | Select-Object TeamName, TeamVisibility, UserName, UserPrincipalName, Role, GroupId |
    Export-Csv -Path "C:\temp\AllTeamMembers.csv" -NoTypeInformation -Encoding UTF8

Write-Host "Export complete. File saved to C:\temp\AllTeamMembers.csv" -ForegroundColor Green
