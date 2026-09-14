# Requires the ImportExcel module. Install it if you haven't already.
Install-Module -Name ImportExcel -Scope CurrentUser

# Path to the Excel file
$excelPath = "C:\Temp\TopLevelACLsImportNEWMonday4-15.xlsx"

# Import the Excel file
$data = Import-Excel -Path $excelPath

# Group data by path
$groupedData = $data | Group-Object -Property Path

$jobs = @()

foreach ($group in $groupedData) {
    Write-Host "Preparing job for path: $($group.Name)"
    $job = Start-Job -ScriptBlock {
        param ($group)
        $acl = Get-Acl -Path $group.Name
        $changesApplied = $false

        foreach ($row in $group.Group) {
            Write-Host "Processing $($row.Identity) for $($group.Name)"
            if ([string]::IsNullOrWhiteSpace($row.Identity) -or [string]::IsNullOrWhiteSpace($row.Path)) {
                Write-Warning "Invalid row with null or empty Identity or Path"
                continue
            }

            $identityRef = New-Object System.Security.Principal.NTAccount($row.Identity)
            $accessRights = $null

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
                $accessRule = New-Object System.Security.AccessControl.FileSystemAccessRule($identityRef, $accessRights, "ContainerInherit,ObjectInherit", "None", [System.Security.AccessControl.AccessControlType]::Allow)
                $acl.AddAccessRule($accessRule)
                $changesApplied = $true
                Write-Host "Added $($accessRights) for $($row.Identity) to $($group.Name)"
            }
        }

        if ($changesApplied) {
            try {
                Set-Acl -Path $group.Name -AclObject $acl
                Write-Host "Successfully set ACL for $($group.Name)"
            } catch {
                Write-Error "Failed to set ACL for $($group.Name): $_"
            }
        }
    } -ArgumentList $group

    $jobs += $job
}

# Wait for all jobs to complete
$jobs | Wait-Job

# Collect results from each job
$jobs | ForEach-Object {
    Receive-Job -Job $_
    Remove-Job -Job $_
}
