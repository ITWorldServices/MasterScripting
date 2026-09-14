# Connect to the Exchange server
$ExchangeServer = "mail-al2.brifutelectric.com" # Replace with your Exchange server name
$Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri http://$ExchangeServer/PowerShell/ -Authentication Kerberos
Import-PSSession $Session -AllowClobber -DisableNameChecking

# Define the output file path
$OutputFile = "C:\temp\ExchangeDistributionGroups.csv" # Replace with the desired output file path

# Create an array to store the distribution group details
$DistributionGroups = @()

# Get all distribution groups
$Groups = Get-DistributionGroup -ResultSize Unlimited

# Iterate through each distribution group and retrieve the details
foreach ($Group in $Groups) {
    $GroupId = $Group.Identity

    # Get the group owners
    $Owners = Get-RecipientPermission -Identity $GroupId |
              Where-Object { $_.Trustee -ne "NT AUTHORITY\SELF" -and $_.AccessRights -eq "FullAccess" } |
              Select-Object -ExpandProperty Trustee

    # Get the group members
    $Members = Get-DistributionGroupMember -Identity $GroupId |
               Select-Object -ExpandProperty Identity

    # Get the group permissions
    $Permissions = Get-ADPermission -Identity $GroupId |
                   Where-Object { $_.User -notlike "NT AUTHORITY\SELF" } |
                   Select-Object -Property @{Name = "Group"; Expression = { $GroupId }},
                                              User,
                                              AccessRights

    # Create a custom object with distribution group details
    $GroupDetails = [PSCustomObject]@{
        Group = $GroupId
        Owners = $Owners
        Members = $Members
        Permissions = $Permissions
    }

    # Add the distribution group details to the array
    $DistributionGroups += $GroupDetails
}

# Export the distribution group details to a CSV file
$DistributionGroups | Export-Csv -Path $OutputFile -NoTypeInformation

# Disconnect from the Exchange server
Remove-PSSession $Session
