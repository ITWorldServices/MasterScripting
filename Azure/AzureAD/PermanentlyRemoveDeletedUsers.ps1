# This script will permanently delete users from the Deleted Users Folder using DirectoryObjectIds.
# You can get these Ids from Azure (Entra ID).
# There will be an export of the users deleted located in C:\Temp and file name time stamped.

# Connect to GCC High MgGraph
# Connect to Microsoft Graph (ensure you have the necessary permissions)
Connect-MgGraph -Environment USGovDoD -Scopes "Directory.ReadWrite.All"

# Connect to Standard MgGraph
# Connect to Microsoft Graph (ensure you have the necessary permissions)
# Connect-MgGraph -Scopes "Directory.ReadWrite.All"

function Is-ValidGuid {
    param ([string]$StringGuid)
    $guidRegex = '^\{?[0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12}\}?$'
    return $StringGuid -match $guidRegex
}

# Array to store deleted users
$deletedUsers = @()

do {
    do {
        $directoryObjectId = Read-Host "Enter the DirectoryObjectId of the item to be permanently deleted"
        if (-not (Is-ValidGuid $directoryObjectId)) {
            Write-Host "Invalid DirectoryObjectId format. Please enter a valid GUID." -ForegroundColor Red
        }
    } while (-not (Is-ValidGuid $directoryObjectId))

    $confirmation = Read-Host "Are you sure you want to permanently delete this item? (Y/N)"
    if ($confirmation -eq 'Y' -or $confirmation -eq 'y') {
        try {
            # Check if the item exists before attempting to delete
            $item = Get-MgDirectoryDeletedItem -DirectoryObjectId $directoryObjectId -ErrorAction Stop
            if ($item) {
                Remove-MgDirectoryDeletedItem -DirectoryObjectId $directoryObjectId -ErrorAction Stop
                Write-Host "Item permanently deleted." -ForegroundColor Green
                
                # Add the deleted item to the array
                $deletedUsers += [PSCustomObject]@{
                    DirectoryObjectId = $directoryObjectId
                    DisplayName = $item.AdditionalProperties.displayName
                    UserPrincipalName = $item.AdditionalProperties.userPrincipalName
                    DeletedDate = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
                }
            } else {
                Write-Host "Item not found in the deleted items. It may have already been removed or never existed." -ForegroundColor Yellow
            }
        }
        catch {
            if ($_.Exception.Message -like "*Request_ResourceNotFound*") {
                Write-Host "Item not found. It may have already been removed or never existed." -ForegroundColor Yellow
            } else {
                Write-Host "An error occurred: $_" -ForegroundColor Red
            }
        }
    } else {
        Write-Host "Deletion cancelled for this item." -ForegroundColor Yellow
    }

    do {
        $continue = Read-Host "Do you want to remove another item? (Y/N)"
        if ($continue -notmatch '^[YyNn]$') {
            Write-Host "Invalid input. Please enter Y or N." -ForegroundColor Red
        }
    } while ($continue -notmatch '^[YyNn]$')

} while ($continue -eq 'Y' -or $continue -eq 'y')

# Export deleted users to CSV
if ($deletedUsers.Count -gt 0) {
    $exportPath = "C:\Temp\PermanentlyDeletedUsers_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv"
    $deletedUsers | Export-Csv -Path $exportPath -NoTypeInformation
    Write-Host "Exported permanently deleted users to: $exportPath" -ForegroundColor Cyan
} else {
    Write-Host "No users were permanently deleted during this session." -ForegroundColor Yellow
}

Write-Host "Script execution completed." -ForegroundColor Cyan