# Define the root path to collect the top-level folders from
$RootPath = "Y:\ttbntfile1\shared\appsdata\groups"

# Ensure the path is correct and accessible
if (-not (Test-Path -Path $RootPath)) {
    Write-Host "The path $RootPath does not exist or is not accessible. Please check the path and permissions."
    exit
}

# Collect the top-level folders from the root path
$folders = Get-ChildItem -Path $RootPath -Directory

# Ensure the output directory exists
$outputDir = "C:\temp"
if (-not (Test-Path -Path $outputDir)) {
    New-Item -ItemType Directory -Path $outputDir
}

# Initialize an array to hold the ACL information for all top-level folders
$AclInfo = @()

foreach ($folder in $folders) {
    $Acl = Get-Acl -Path $folder.FullName
    foreach ($AccessRule in $Acl.Access) {
        $AclInfo += [PSCustomObject]@{
            Path        = $folder.FullName
            Identity    = $AccessRule.IdentityReference.ToString()
            AccessType  = $AccessRule.AccessControlType
            Rights      = $AccessRule.FileSystemRights
            Inheritance = $AccessRule.IsInherited
        }
    }
}

# Export the ACL information of top-level folders to a CSV file
$outputFile = Join-Path $outputDir "TopLevelFoldersACLs.csv"
if ($AclInfo.Count -gt 0) {
    $AclInfo | Export-Csv -Path $outputFile -NoTypeInformation -Force
}
