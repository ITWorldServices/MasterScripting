# Generate a 14-character random password with uppercase, lowercase, numbers, and special characters
function Generate-RandomPassword {
    $Length = 14
    $UpperCase = [char[]](65..90) # A-Z
    $LowerCase = [char[]](97..122) # a-z
    $Numbers = [char[]](48..57) # 0-9
    $SpecialChars = @('!', '@', '#', '$', '%', '^', '&', '*', '(', ')', '-', '_', '=', '+')

    # Ensure at least one character from each set
    $Password = @(
        ($UpperCase | Get-Random -Count 1),
        ($LowerCase | Get-Random -Count 1),
        ($Numbers | Get-Random -Count 1),
        ($SpecialChars | Get-Random -Count 1)
    )

    # Fill the rest of the password length with random characters from all sets
    $AllChars = $UpperCase + $LowerCase + $Numbers + $SpecialChars
    $Password += ($AllChars | Get-Random -Count ($Length - 4))

    # Shuffle the password array to randomize character order and convert to string
    -join ($Password | Get-Random -Count $Length)
}

# Store the generated password in a secure string
$RandomPassword = Generate-RandomPassword
$SecurePassword = ConvertTo-SecureString -String $RandomPassword -AsPlainText -Force

# Create the local admin account named 'rs-local'
New-LocalUser -Name "rs-local" -Password $SecurePassword -FullName "Local Admin Account" -Description "Created via Intune Autodeploy" -PasswordNeverExpires:$true

# Add the account to the Administrators group
Add-LocalGroupMember -Group "Administrators" -Member "rs-local"

# Output the generated password for logging or debugging purposes (optional)
Write-Host "Generated Password: $RandomPassword"
