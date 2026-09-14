#Calculate the disk space used by each share on a server
# This is a short powershell script that will collect the names of the shares on a computer
# Remove the default share names. And then calculate the storage used by the remaining share names.
#  *** Depending on the size of the share this could take a very long time!!!
# Created by: DPH 

# Get a list of all shares on the local machine
$shares = Get-SmbShare | Select-Object -ExpandProperty Name

# Define an array of default Windows server shares
$defaultShares = @('ADMIN$', 'C$', 'INC$', 'IPC$', 'NETLOGON', 'SYSVOL', 'print$', 'Netwrix_Auditor_Subscriptions$', 'Netwrix_Auditor_Subscriptions$', 'Netwrix_UAVR$')

# Filter out the default shares from the list
$filteredShares = $shares | Where-Object { $defaultShares -notcontains $_ }

# Loop through each share in the list and calculate its size
foreach ($sharename in $filteredShares) {
    Get-SmbShare -Name $sharename | Select-Object Name, Path, @{Name="SizeOnDisk"; Expression={"{0:N2} GB" -f ((Get-ChildItem $_.Path -Recurse | Measure-Object Length -Sum).Sum / 1GB)}}
}

