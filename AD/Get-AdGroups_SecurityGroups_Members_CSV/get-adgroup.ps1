# Import the Active Directory module
Import-Module ActiveDirectory

# Get all security groups
$securityGroups = Get-ADGroup -Filter {GroupCategory -eq 'Security'}

# Create an array to store group information
$groupInfo = @()

# Loop through each security group
foreach ($group in $securityGroups) {
    # Get group members
    $members = Get-ADGroupMember -Identity $group | Select-Object Name, SamAccountName

    # Create an object for each group and its members
    foreach ($member in $members) {
        $groupInfo += [PSCustomObject]@{
            "GroupName" = $group.Name
            "GroupSamAccountName" = $group.SamAccountName
            "MemberName" = $member.Name
            "MemberSamAccountName" = $member.SamAccountName
        }
    }
}

# Export the group information to a CSV file
$groupInfo | Export-Csv -Path "SecurityGroupsAndMembers.csv" -NoTypeInformation