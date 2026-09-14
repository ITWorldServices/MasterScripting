# Define the list of folders to check permissions for
$folders = @(
    "Y:\ttbntfile1\shared\appsdata\groups\0-Policies and Procedures Manual"

# Add additional directories below the first line. Separate each by a comma at the end of the line after the " mark. 
)

# Ensure the output directory exists
$outputDir = "C:\temp"
if (-not (Test-Path -Path $outputDir)) {
    New-Item -ItemType Directory -Path $outputDir
}

foreach ($folder in $folders) {
    $folderName = Split-Path $folder -Leaf
    $safeFolderName = $folderName -replace ' ', ''
    $outputFile = Join-Path $outputDir "$safeFolderName.csv"

    # Initialize an array to hold the ACL information
    $AclInfo = @()

    # Recursively get all items in the specified directory, including directories
    $Items = Get-ChildItem -Path $folder -Recurse -Directory -Force -ErrorAction SilentlyContinue

    if ($Items.Count -eq 0) {
        Write-Host "No subfolders found in $folder."
    } else {
        foreach ($Item in $Items) {
            try {
                $Acl = Get-Acl -Path $Item.FullName
                foreach ($Access in $Acl.Access) {
                    $AclInfo += [PSCustomObject]@{
                        Path        = $Item.FullName
                        Identity    = $Access.IdentityReference.ToString()
                        AccessType  = $Access.AccessControlType
                        Rights      = $Access.FileSystemRights
                        Inheritance = $Access.IsInherited
                    }
                }
            } catch {
                Write-Host "Failed to get ACL for $($Item.FullName): $_"
            }
        }

        # Export the ACL information to a CSV file, if any was found
        if ($AclInfo.Count -gt 0) {
            $AclInfo | Export-Csv -Path $outputFile -NoTypeInformation -Force
            Write-Host "ACLs exported to $outputFile"
        } else {
            Write-Host "ACL information was not found for any items in $folder."
        }
    }
}
