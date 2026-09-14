#Define Transcript Log Path and file name
$transcriptpath = "C\Temp\Transcripts\RenameAllDisabledUsersADObjectsBulk.txt"


# Define the path to the CSV file
$csvPath = "C:\Temp\DisableAllAccountsFirstLastName.csv"

Start-Transcript -Path $transcriptpath

# Import the CSV file
$users = Import-Csv -Path $csvPath

# Initialize an array to hold the DistinguishedNames
$distinguishedNames = @()

# Loop through each user in the CSV
foreach ($user in $users) {
    # Define variables from the CSV headers
    $SamAccount = $user.SamAccountName
    $NewDisplayName = $user.NewDisplayName
    $DisplayName = $user.DisplayName

    # Retrieve the DistinguishedName using Get-ADUser
    $DistinguishedName = (Get-ADUser -Filter {SamAccountName -eq $SamAccount}).DistinguishedName

    # Add the DistinguishedName to the array
    $distinguishedNames += $DistinguishedName

    
    # Run the Rename-ADObject command
    try {
        Rename-ADObject -Identity $DistinguishedName -NewName $NewDisplayName
        Write-Host "Successfully renamed $DisplayName to $NewDisplayName" -ForegroundColor Green
        } catch {
        Write-Host "Failed to rename $DisplayName to $NewDisplayName" -ForegroundColor Red
        }
        } 


# Output the array of DistinguishedNames to verify all users renamed
$distinguishedNames

Stop-Transcript