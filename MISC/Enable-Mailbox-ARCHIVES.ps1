# Connect to Exchange Online

# Import the CSV file
$users = Import-Csv -Path C:\temp\tandBArchives.csv

# Loop through each user in the CSV and enable the archive
foreach ($user in $users) {
    Enable-Mailbox -Identity $user.Email -Archive
}

# Disconnect the session