# Requires the ImportExcel module. Install it if you haven't already.
Install-Module -Name ImportExcel -Scope CurrentUser

# Path to the Excel file
$excelPath = "C:\Temp\0-PoliciesandProceduresManual.xlsx"

# Import the Excel file
$data = Import-Excel -Path $excelPath

foreach ($row in $data) {
    if ($null -eq $row.Identity) {
        Write-Warning "Identity is null for path $($row.Path)"
        continue
    }

    if ($null -eq $row.Path) {
        Write-Warning "Path is null for identity $($row.Identity)"
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

