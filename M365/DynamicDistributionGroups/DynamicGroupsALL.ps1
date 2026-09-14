# Connect to the Exchange server
# $ExchangeServer = "MAIL-AL2.brifutelectric.com" # Replace with your Exchange server name
# $Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri http://$ExchangeServer/PowerShell/ -Authentication Kerberos
# Import-PSSession $Session -AllowClobber -DisableNameChecking

# Define the output file path
$OutputFile = "C:\temp\ExchangeDynamicGroups.csv" # Replace with the desired output file path

# Create an array to store the dynamic distribution group details
$DynamicGroups = @()

# Get all dynamic distribution groups
$Groups = Get-DynamicDistributionGroup -ResultSize Unlimited

# Iterate through each dynamic distribution group and retrieve the details
foreach ($Group in $Groups) {
    $GroupId = $Group.Identity

    # Get the group owners
    $Owners = Get-RecipientPermission -Identity $GroupId |
              Where-Object { $_.Trustee -ne "NT AUTHORITY\SELF" -and $_.AccessRights -eq "FullAccess" } |
              Select-Object -ExpandProperty Trustee

    # Get the group members
    $Members = Get-Recipient -RecipientPreviewFilter $Group.RecipientFilter |
               Select-Object -ExpandProperty Identity

    # Get the group permissions
    $Permissions = Get-ADPermission -Identity $GroupId |
                   Where-Object { $_.User -notlike "NT AUTHORITY\SELF" } |
                   Select-Object -Property @{Name = "Group"; Expression = { $GroupId }},
                                              User,
                                              AccessRights

    # Create a custom object with dynamic distribution group details
    $GroupDetails = [PSCustomObject]@{
        Group = $GroupId
        Owners = $Owners
        Members = $Members
        Permissions = $Permissions
    }

    # Add the dynamic distribution group details to the array
    $DynamicGroups += $GroupDetails
}

# Export the dynamic distribution group details to a CSV file
$DynamicGroups | Export-Csv -Path $OutputFile -NoTypeInformation

# Disconnect from the Exchange server
# Remove-PSSession $Session
