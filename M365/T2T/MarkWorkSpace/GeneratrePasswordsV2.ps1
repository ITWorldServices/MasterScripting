<#
.SYNOPSIS
This PowerShell script generates a specified number of strong passwords and exports them to a CSV file.

.DESCRIPTION
The script prompts the user for:
1. The number of passwords to generate.
2. An identifier to prepend to the output file name.
3. The length of the passwords.
4. Options to include uppercase, lowercase, numbers, and special characters.

It generates strong passwords based on the chosen character sets and saves them in a CSV file with a dynamic name that includes the provided identifier and a timestamp.
The script ensures user-friendly interaction by clearing the console at the start and validating user inputs.
#>

# Clear the screen
Clear-Host

# Function to get yes/no input from the user
function Get-YesNoInput {
    param (
        [string]$Prompt
    )
    do {
        $inputValue = Read-Host "$Prompt (yes/no, y/n)"
        $inputValue = $inputValue.Trim().ToLower()
        if ($inputValue -notin @('yes', 'y', 'no', 'n')) {
            Write-Host "Invalid input. Please enter yes, y, no, or n." -ForegroundColor Red
        }
    } until ($inputValue -in @('yes', 'y', 'no', 'n'))
    return $inputValue -in @('yes', 'y')
}

# Prompt the user for the number of passwords to generate
do {
    $NumberOfPasswords = Read-Host "Enter the number of passwords to generate"
    if (-not [int]::TryParse($NumberOfPasswords, [ref]$null) -or [int]$NumberOfPasswords -le 0) {
        Write-Host "Invalid input. Please enter a positive integer." -ForegroundColor Red
    }
} until ([int]::TryParse($NumberOfPasswords, [ref]$null) -and [int]$NumberOfPasswords -gt 0)
$NumberOfPasswords = [int]$NumberOfPasswords

# Prompt the user for an identifier to prepend to the file name
do {
    $FileIdentifier = Read-Host "Enter an identifier to prepend to the file name"
    if ([string]::IsNullOrWhiteSpace($FileIdentifier)) {
        Write-Host "Invalid input. Please enter a valid identifier." -ForegroundColor Red
    }
} until (-not [string]::IsNullOrWhiteSpace($FileIdentifier))

# Prompt the user for the password length
do {
    $PasswordLength = Read-Host "Enter the length of the passwords (default is 16)"
    if (-not [int]::TryParse($PasswordLength, [ref]$null)) {
        $PasswordLength = 16
        Write-Host "Invalid input. Using default length of 16." -ForegroundColor Yellow
    } elseif ([int]$PasswordLength -le 0) {
        Write-Host "Invalid input. Please enter a positive integer." -ForegroundColor Red
    }
} until ([int]::TryParse($PasswordLength, [ref]$null) -and [int]$PasswordLength -gt 0)
$PasswordLength = [int]$PasswordLength

# Prompt the user for character set options
$IncludeUppercase = Get-YesNoInput "Include uppercase letters?"
$IncludeLowercase = Get-YesNoInput "Include lowercase letters?"
$IncludeNumbers = Get-YesNoInput "Include numbers?"
$IncludeSpecialChars = Get-YesNoInput "Include special characters?"

# Ensure at least one character type is selected
while (-not ($IncludeUppercase -or $IncludeLowercase -or $IncludeNumbers -or $IncludeSpecialChars)) {
    Write-Host "Please select at least one character type." -ForegroundColor Red
    $IncludeUppercase = Get-YesNoInput "Include uppercase letters?"
    $IncludeLowercase = Get-YesNoInput "Include lowercase letters?"
    $IncludeNumbers = Get-YesNoInput "Include numbers?"
    $IncludeSpecialChars = Get-YesNoInput "Include special characters?"
}

# Get current date and time in a specific format
$CurrentDateTime = Get-Date -Format "yyyyMMdd_HHmmss"

# Construct the output file path with identifier and timestamp
$OutputFilePath = "C:\temp\$FileIdentifier`_Passwords_$CurrentDateTime.csv"

# Check if output directory exists
if (-not (Test-Path -Path "C:\temp")) {
    Write-Host "Output directory does not exist. Creating it..." -ForegroundColor Yellow
    New-Item -Path "C:\temp" -ItemType Directory
}

# Function to generate a strong password
function Generate-StrongPassword {
    param (
        [int]$Length,
        [bool]$IncludeUppercase,
        [bool]$IncludeLowercase,
        [bool]$IncludeNumbers,
        [bool]$IncludeSpecialChars
    )

    $chars = @()
    if ($IncludeUppercase) { $chars += [char[]]"ABCDEFGHJKLMNPQRSTUVWXYZ" }
    if ($IncludeLowercase) { $chars += [char[]]"abcdefghijkmnpqrstuvwxyz" }
    if ($IncludeNumbers) { $chars += [char[]]"23456789" }
    if ($IncludeSpecialChars) { $chars += [char[]]"!#$%^&-;:<>?" }

    # Ensure password includes at least one of each selected character type
    $password = @()
    if ($IncludeUppercase) { $password += Get-Random -InputObject ([char[]]"ABCDEFGHJKLMNPQRSTUVWXYZ") }
    if ($IncludeLowercase) { $password += Get-Random -InputObject ([char[]]"abcdefghijkmnpqrstuvwxyz") }
    if ($IncludeNumbers) { $password += Get-Random -InputObject ([char[]]"23456789") }
    if ($IncludeSpecialChars) { $password += Get-Random -InputObject ([char[]]"!#$%^&-;:<>?") }

    # Fill the rest of the password length with random characters from the selected sets
    for ($i = $password.Count; $i -lt $Length; $i++) {
        $password += Get-Random -InputObject $chars
    }

    # Shuffle the array to avoid the first characters always being from specific sets
    $password = $password | Get-Random -Count $password.Count

    -join $password
}

# Create an array to store the passwords
$passwords = @()

# Generate the required number of passwords
for ($i = 1; $i -le $NumberOfPasswords; $i++) {
    $password = Generate-StrongPassword -Length $PasswordLength -IncludeUppercase $IncludeUppercase -IncludeLowercase $IncludeLowercase -IncludeNumbers $IncludeNumbers -IncludeSpecialChars $IncludeSpecialChars
    $passwords += [PSCustomObject]@{
        UserPrincipalName = "UPN$i"   # Placeholder user identifier (e.g., User1, User2)
        Password   = $password
    }
}

# Export the passwords to a CSV file
try {
    $passwords | Export-Csv -Path $OutputFilePath -NoTypeInformation -Encoding UTF8
    Write-Host "Passwords generated and saved to: $OutputFilePath" -ForegroundColor Green
} catch {
    Write-Host "Error exporting passwords to CSV: $($Error[0].Message)" -ForegroundColor Red
}

Write-Host "Generated $NumberOfPasswords passwords." -ForegroundColor Green
