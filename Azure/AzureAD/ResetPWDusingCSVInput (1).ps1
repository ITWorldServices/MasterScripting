# Connect to Office 365
Connect-AzureAD

# Define the path to the CSV file
$CsvFilePath = "C:\temp\DMS-Source-UPNandPWD.csv"

# Read user information from the CSV file, colunms will include UserPrincipalName and Password
$Users = Import-Csv -Path $CsvFilePath

# Loop through each user in the CSV file
foreach ($User in $Users) {
    $UserPrincipalName = $User.UserPrincipalName
    $Password = $User.Password

    # Create the password profile
    $PasswordProfile = @{
        "Password" = $Password
        "ForceChangePasswordNextLogin" = $false  # Set to $true if you want the user to change the password at the next login
    }

    # Set the user's password profile
    Set-AzureADUser -ObjectId $UserPrincipalName -PasswordProfile $PasswordProfile

    # Display information about the password profile for the user
    Write-Host "Password profile set for $UserPrincipalName"
    Write-Host "  Password: $Password"
    Write-Host "  ForceChangePasswordNextLogin: $($PasswordProfile.ForceChangePasswordNextLogin)"
}
