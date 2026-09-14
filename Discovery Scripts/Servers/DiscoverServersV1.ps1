# This script creates a list of all devices in the domain with a Server operating system installed.

# Import the Active Directory module
Import-Module ActiveDirectory

# Calculate the cutoff date (30 days ago)
$cutoffDate = (Get-Date).AddDays(-30)

# Get all computer objects from Active Directory that have been active in the last 30 days and have a server name containing "Windows Server"
$computers = Get-ADComputer -Filter {LastLogonTimeStamp -gt $cutoffDate -and OperatingSystem -like "*Windows Server*"} -Property *

# Set the output path
$outputPath = ".\serverlist.csv"

# Export the computer objects to a CSV file with modified column header
$computers | Select-Object @{Name="Servername"; Expression={$_.Name}} | Export-Csv -Path $outputPath -NoTypeInformation

# Provide output confirmation
Write-Host "CSV file exported to $outputPath"
