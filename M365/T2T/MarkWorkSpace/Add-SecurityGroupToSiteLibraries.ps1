# Previous/default values
# $siteUrl = $siteUrlDefault = "https://blackhawkcenter.sharepoint.com/teams/GoodwinAustinTeam"
# $groupEmail = $groupEmailDefault = "GoodwinAustinTeamAccess@goodwintx.com"
$permissionLevel = $permissionLevelDefault = "Contribute"
# $clientId = $clientIdDefault = "5e940a58-2d36-4e9e-9d6f-47009e181aea"

# Prompt for Site URL
$input = Read-Host "Enter SharePoint Site URL [$siteUrlDefault]"
if (![string]::IsNullOrWhiteSpace($input)) { $siteUrl = $input }

# Prompt for Group Email
$input = Read-Host "Enter Group Email [$groupEmailDefault]"
if (![string]::IsNullOrWhiteSpace($input)) { $groupEmail = $input }

# Prompt for Permission Level
Write-Host "Select Permission Level:"
Write-Host "1 - Read"
Write-Host "2 - Contribute"
Write-Host "3 - Edit"
Write-Host "4 - Full Control"
$input = Read-Host "Enter number for permission level [2]"
switch ($input) {
    "1" { $permissionLevel = "Read" }
    "2" { $permissionLevel = "Contribute" }
    "3" { $permissionLevel = "Edit" }
    "4" { $permissionLevel = "Full Control" }
    default { $permissionLevel = $permissionLevelDefault }
}

# Prompt for Client ID
$input = Read-Host "Enter Client ID [$clientIdDefault]"
if (![string]::IsNullOrWhiteSpace($input)) { $clientId = $input }

# Connect to SharePoint
Connect-PnPOnline -Url $siteUrl -ClientId $clientId -Interactive

# Get document libraries except Style Library
$libraries = Get-PnPList | Where-Object { $_.BaseTemplate -eq 101 -and $_.Title -ne "Style Library" }

# Display libraries
Write-Host "Document Libraries to update:"
foreach ($lib in $libraries) {
    Write-Host " - $($lib.Title)"
}

# Prompt to continue
$continue = Read-Host "Continue and update permissions? (Y/N)"
if ($continue -notin @("Y", "y")) {
    Write-Host "Operation cancelled."
    exit
}

foreach ($lib in $libraries) {
    try {
        Write-Host "Processing $($lib.Title)..."
        Set-PnPList -Identity $lib -BreakRoleInheritance -ErrorAction Stop
        Set-PnPListPermission -Identity $lib -User $groupEmail -AddRole $permissionLevel -ErrorAction Stop
        Write-Host "Permissions set for $($lib.Title)."
    } catch {
        Write-Warning "Failed for $($lib.Title): $_"
    }
}
