# Connect to Exchange Online
Connect-ExchangeOnline
# Set the path for the CSV file to save the exported group and member data
$exportPath = "C:\temp\FESO365GroupMembers.csv"
# Get all Office 365 groups
$groups = Get-UnifiedGroup -ResultSize Unlimited
# Create an array to store group and member data
$groupAndMemberDataArray = @()
# Loop through each group and collect information about members
foreach ($group in $groups) {
    $groupData = @{
        "GroupName" = $group.DisplayName
        "GroupEmail" = $group.PrimarySmtpAddress
        "GroupMembers" = (Get-UnifiedGroupLinks -Identity $group.DistinguishedName -LinkType Members).Name -join ', '
    }
    $groupAndMemberDataArray += New-Object PSObject -Property $groupData
}
# Export group and member data to CSV
$groupAndMemberDataArray | Export-Csv -Path $exportPath -NoTypeInformation
# Disconnect from Exchange Online
Disconnect-ExchangeOnline