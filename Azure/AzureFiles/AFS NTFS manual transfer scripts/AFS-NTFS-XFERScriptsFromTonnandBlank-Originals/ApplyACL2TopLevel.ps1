# Requires the ImportExcel module. Install it if you haven't already.
Install-Module -Name ImportExcel -Scope CurrentUser

# Path to the Excel file
$excelPath = "C:\Temp\TopLevelACLsImport.xlsx"


# Import the Excel file
$data = Import-Excel -Path $excelPath

foreach ($row in $data) {
    # Check if the Identity or Path is null or empty
    if ([string]::IsNullOrWhiteSpace($row.Identity)) {
        Write-Warning "Identity is null or empty for path $($row.Path)"
        continue
    }

    if ([string]::IsNullOrWhiteSpace($row.Path)) {
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
            "Modify, ChangePermissions, Synchronize" {
                $accessRights = [System.Security.AccessControl.FileSystemRights]::Modify -bor [System.Security.AccessControl.FileSystemRights]::ChangePermissions -bor [System.Security.AccessControl.FileSystemRights]::Synchronize
            }
            "Modify, Synchronize" {
                $accessRights = [System.Security.AccessControl.FileSystemRights]::Modify -bor [System.Security.AccessControl.FileSystemRights]::Synchronize
            }
            "ReadAndExecute, Synchronize" {
                $accessRights = [System.Security.AccessControl.FileSystemRights]::ReadAndExecute -bor [System.Security.AccessControl.FileSystemRights]::Synchronize
            }
            "Read, Synchronize" {
                $accessRights = [System.Security.AccessControl.FileSystemRights]::Read -bor [System.Security.AccessControl.FileSystemRights]::Synchronize
            }
            "ExecuteFile, Synchronize" {
                $accessRights = [System.Security.AccessControl.FileSystemRights]::ExecuteFile -bor [System.Security.AccessControl.FileSystemRights]::Synchronize
            }
            "Write, ReadAndExecute, Synchronize" {
                $accessRights = [System.Security.AccessControl.FileSystemRights]::Write -bor [System.Security.AccessControl.FileSystemRights]::ReadAndExecute -bor [System.Security.AccessControl.FileSystemRights]::Synchronize
            }
            "Write, Synchronize" {
                $accessRights = [System.Security.AccessControl.FileSystemRights]::Write -bor [System.Security.AccessControl.FileSystemRights]::Synchronize
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
