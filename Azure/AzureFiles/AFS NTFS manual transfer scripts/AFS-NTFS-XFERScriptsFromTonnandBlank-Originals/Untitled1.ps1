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

# Process each top-level folder in a separate job
foreach ($folder in $folders) {
    Start-Job -ScriptBlock {
        param($folderPath, $outputDir)
        $folderName = Split-Path -Path $folderPath -Leaf
        $safeFolderName = $folderName -replace ' ', ''
        $outputFile = Join-Path $outputDir "$safeFolderName.csv"

        $AclInfo = @()
        # Get only subfolders, not files
        $Items = Get-ChildItem -Path $folderPath -Recurse -Directory -Force -ErrorAction SilentlyContinue

        foreach ($Item in $Items) {
            try {
                $Acl = Get-Acl -Path $Item.FullName
                foreach ($accessRule in $Acl.Access) {
                    $AclInfo += [PSCustomObject]@{
                        Path        = $Item.FullName
                        Identity    = $accessRule.IdentityReference.ToString()
                        AccessType  = $accessRule.AccessControlType -as [string]
                        Rights      = $accessRule.FileSystemRights -as [string]
                        Inheritance = $accessRule.IsInherited -as [bool]
                    }
                }
            } catch {
                Write-Host "Failed to get ACL for $($Item.FullName): $_"
            }
        }

        # Export the ACL information of folders to a CSV file
        if ($AclInfo.Count -gt 0) {
            $AclInfo | Export-Csv -Path $outputFile -NoTypeInformation -Force
        }
    } -ArgumentList $folder.FullName, $outputDir
}

# Wait for all jobs to complete
Get-Job | Wait-Job

# Cleanup - remove completed jobs
Get-Job | Remove-Job
