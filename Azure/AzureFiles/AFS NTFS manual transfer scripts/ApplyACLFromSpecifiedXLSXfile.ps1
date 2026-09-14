# Requires the ImportExcel module. Install it if you haven't already.
Install-Module -Name ImportExcel -Scope CurrentUser

# Path to the Excel file
$excelPath = "C:\Temp\SpecificFileName.xlsx"

# Import the Excel file
$data = Import-Excel -Path $excelPath

# If there is an issue with the Excel structure, the $data variable might be empty
if ($data -eq $null) {
    Write-Error "No data could be imported from the Excel file. Check the file path and structure."
    exit
}

foreach ($row in $data) {
    # Adding more detailed output to see what is actually contained in each row
    Write-Host "Reading row: Path: '$($row.Path)', Identity: '$($row.Identity)', AccessType: '$($row.AccessType)', Rights: '$($row.Rights)', Inheritance: '$($row.Inheritance)'"
    
    if ($null -eq $row.Identity -or $row.Identity -eq '') {
        Write-Warning "Identity is null or empty for path $($row.Path)"
        continue
    }

    if ($null -eq $row.Path) {
        Write-Warning "Path is null or empty for identity $($row.Identity)"
        continue
    }

    if ($row.Inheritance -ne "TRUE") {
        $identityRef = New-Object System.Security.Principal.NTAccount($row.Identity)
        $accessRights = $null
        
        # Check for the correct rights based on the entry in the Excel file
        switch -Regex ($row.Rights) {
            "FullControl" {
                $accessRights = [System.Security.AccessControl.FileSystemRights]::FullControl
            }
            "ReadAndExecute, Synchronize" {
                $accessRights = [System.Security.AccessControl.FileSystemRights]::ReadAndExecute -bor [System.Security.AccessControl.FileSystemRights]::Synchronize
            }
            "Modify, Synchronize" {
                $accessRights = [System.Security.AccessControl.FileSystemRights]::Modify -bor [System.Security.AccessControl.FileSystemRights]::Synchronize
            }
            default {
                Write-Warning "Unrecognized rights: $($row.Rights)"
                continue
            }
        }

        if ($accessRights -ne $null) {
            $accessControlType = [System.Security.AccessControl.AccessControlType]::Allow

            $accessRule = New-Object System.Security.AccessControl.FileSystemAccessRule($identityRef, $accessRights, "ContainerInherit,ObjectInherit", "None", $accessControlType)

            try {
                $acl = Get-Acl -Path $row.Path
                $acl.AddAccessRule($accessRule)
                Set-Acl -Path $row.Path -AclObject $acl
            } catch {
                Write-Error "Error applying ACL to path $($row.Path): $_"
            }
        } else {
            Write-Warning "Invalid or unrecognized rights for $($row.Identity) on path $($row.Path)"
        }
    }
}
