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

    # Only process if Inheritance is not set to TRUE
    if ($row.Inheritance -ne "TRUE") {
        # Define permissions string for icacls
        $permissions = $null
        switch -Regex ($row.Rights) {
            "FullControl" {
                $permissions = ":(OI)(CI)F"
            }
            "Modify, ChangePermissions, Synchronize" {
                $permissions = ":(OI)(CI)M"
            }
            "Modify, Synchronize" {
                $permissions = ":(OI)(CI)M"
            }
            "ReadAndExecute, Synchronize" {
                $permissions = ":(OI)(CI)RX"
            }
            "Read, Synchronize" {
                $permissions = ":(OI)(CI)R"
            }
            "ExecuteFile, Synchronize" {
                $permissions = ":(OI)(CI)X"
            }
            "Write, ReadAndExecute, Synchronize" {
                $permissions = ":(OI)(CI)W, RX"
            }
            "Write, Synchronize" {
                $permissions = ":(OI)(CI)W"
            }
            default {
                Write-Warning "Unrecognized rights: $($row.Rights)"
                continue
            }
        }

        if ($permissions -ne $null) {
            # Build the icacls command
            $icaclsCommand = "icacls `"$($row.Path)`" /grant `"$($row.Identity)$permissions`" /T /C"
            Write-Host "Applying permissions to: $($row.Path)"

            try {
                # Execute the icacls command and capture output
                $output = Invoke-Expression $icaclsCommand
                if ($output -ne $null) {
                    $outputLines = $output -split "`r`n"  # Split the output into lines
                   # foreach ($line in $outputLines) {
                   #     Write-Host $line  # Print each line separately
                   # }
                }
            } catch {
                Write-Error "Error applying ACL to path $($row.Path): $_"
            }
        } else {
            Write-Warning "Invalid or unrecognized rights for $($row.Identity) on path $($row.Path)"
        }
    }
}
