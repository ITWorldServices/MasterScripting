# Import necessary functions from Microsoft.Graph
Import-Module Microsoft.Graph.Groups -Function Get-MgGroup, Get-MgGroupMember, Get-MgGroupOwner, Get-MgTeam
Import-Module Microsoft.Graph.Users -Function Get-MgUser

# Connect to Microsoft Graph (this will prompt for authentication)
Connect-MgGraph -Scopes "Group.Read.All", "User.Read.All", "Directory.Read.All"

# Define the properties we want to export
$properties = @(
    "Id", "CreatedDateTime", "Description", "DisplayName", 
    "Mail", "MailEnabled", "MailNickname", "RenewedDateTime", 
    "SecurityEnabled", "SecurityIdentifier", "Visibility"
)

# Get all groups with the specified properties
$groups = Get-MgGroup -Property ($properties + @("GroupTypes")) -All

# Process and export the groups
$exportData = $groups | ForEach-Object {
    $group = $_
    $groupData = $group | Select-Object $properties

    # Join GroupTypes array into a single string
    $groupData | Add-Member -NotePropertyName "GroupTypes" -NotePropertyValue ($group.GroupTypes -join "; ")

    # Get expanded properties
    $members = Get-MgGroupMember -GroupId $group.Id
    $owners = Get-MgGroupOwner -GroupId $group.Id
    $settings = Get-MgGroupSetting -GroupId $group.Id
    $team = Get-MgTeam -GroupId $group.Id -ErrorAction SilentlyContinue

    # Add member and owner information
    $groupData | Add-Member -NotePropertyName "MemberCount" -NotePropertyValue $members.Count
    $groupData | Add-Member -NotePropertyName "OwnerCount" -NotePropertyValue $owners.Count
    
    $memberUPNs = $members | Where-Object { $_.AdditionalProperties.userPrincipalName } | 
                  ForEach-Object { $_.AdditionalProperties.userPrincipalName }
    $ownerUPNs = $owners | Where-Object { $_.AdditionalProperties.userPrincipalName } | 
                 ForEach-Object { $_.AdditionalProperties.userPrincipalName }
    
    $groupData | Add-Member -NotePropertyName "Members" -NotePropertyValue ($memberUPNs -join "; ")
    $groupData | Add-Member -NotePropertyName "Owners" -NotePropertyValue ($ownerUPNs -join "; ")

    # Add settings information
    $groupData | Add-Member -NotePropertyName "SettingsCount" -NotePropertyValue $settings.Count
    $settingsInfo = $settings | ForEach-Object { "$($_.DisplayName): $($_.Values.Name -join ', ')" }
    $groupData | Add-Member -NotePropertyName "Settings" -NotePropertyValue ($settingsInfo -join "; ")

    # Add Team information if available
    if ($team) {
        $groupData | Add-Member -NotePropertyName "IsTeam" -NotePropertyValue $true
        $groupData | Add-Member -NotePropertyName "TeamDisplayName" -NotePropertyValue $team.DisplayName
        $groupData | Add-Member -NotePropertyName "TeamDescription" -NotePropertyValue $team.Description
        $groupData | Add-Member -NotePropertyName "TeamVisibility" -NotePropertyValue $team.Visibility
    } else {
        $groupData | Add-Member -NotePropertyName "IsTeam" -NotePropertyValue $false
    }

    # Add additional properties
    $groupData | Add-Member -NotePropertyName "Classification" -NotePropertyValue $group.Classification
    $groupData | Add-Member -NotePropertyName "CreatedDateTime" -NotePropertyValue $group.CreatedDateTime
    $groupData | Add-Member -NotePropertyName "ExpirationDateTime" -NotePropertyValue $group.ExpirationDateTime
    $groupData | Add-Member -NotePropertyName "IsAssignableToRole" -NotePropertyValue $group.IsAssignableToRole
    $groupData | Add-Member -NotePropertyName "MembershipRule" -NotePropertyValue $group.MembershipRule
    $groupData | Add-Member -NotePropertyName "MembershipRuleProcessingState" -NotePropertyValue $group.MembershipRuleProcessingState
    $groupData | Add-Member -NotePropertyName "PreferredDataLocation" -NotePropertyValue $group.PreferredDataLocation
    $groupData | Add-Member -NotePropertyName "PreferredLanguage" -NotePropertyValue $group.PreferredLanguage
    $groupData | Add-Member -NotePropertyName "Theme" -NotePropertyValue $group.Theme

    $groupData
}

# Export the processed data to a CSV file
$exportData | Export-Csv -Path "C:\Temp\GEM-GroupExport.csv" -NoTypeInformation

Write-Host "Group information has been exported to C:\Temp\GEM-GroupExport.csv"