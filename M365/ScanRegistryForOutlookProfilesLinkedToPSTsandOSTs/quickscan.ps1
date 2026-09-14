# Function to convert a binary value to a string
function ConvertFrom-BinaryToString {
    param (
        [Parameter(Mandatory=$true)]
        [byte[]]$BinaryValue
    )

    return [System.Text.Encoding]::Unicode.GetString($BinaryValue)
}

# Function to remove leading control characters from a string
function RemoveLeadingControlCharacters {
    param (
        [Parameter(Mandatory=$true)]
        [string]$StringValue
    )

    # Remove leading control characters and return the cleaned-up string
    return $StringValue -replace "^[\\u0000-\\u001F\\u007F]+", ""
}

# List of versions to check
$versions = @("11.0", "12.0", "14.0", "15.0", "16.0")

foreach ($version in $versions) {
    # Define the registry path to scan for the current version
    $registryPath = "HKCU:\SOFTWARE\Microsoft\Office\$version\Outlook\Profiles"

    # Check if the registry path exists
    if (Test-Path $registryPath) {
        Write-Host "Searching within registry path: $registryPath"

        # Enumerate all the subkeys
        Get-ChildItem -Path $registryPath -Recurse | ForEach-Object {
            $values = Get-ItemProperty $_.PSPath

            foreach ($property in $values.PSObject.Properties) {
                if ($property.Name -notin @("PSPath", "PSParentPath", "PSChildName", "PSDrive", "PSProvider")) {
                    # Check if value is of type byte[] (Binary)
                    if ($property.Value -is [byte[]]) {
                        $stringValue = ConvertFrom-BinaryToString -BinaryValue $property.Value
                        $cleanedValue = RemoveLeadingControlCharacters -StringValue $stringValue

                        if ($cleanedValue -like "*.pst*" -or $cleanedValue -like "*.ost*") {
                            Write-Host "Path: $($_.Name) - Name: $($property.Name) - Value: $cleanedValue"
                        }
                    }
                }
            }
        }
    } else {
        Write-Host "Registry path not found: $registryPath"
    }
}
