# Connect to Exchange Online

# Import the CSV file
$csvData = Import-Csv -Path "C:\temp\NWI-DLs.csv"

# Loop through each row in the CSV
foreach ($row in $csvData) {
    # Create a new distribution group
    New-DistributionGroup -Name $row.Name -Alias $row.Alias -PrimarySmtpAddress $row.Email
}

# Disconnect the session