<#
.SYNOPSIS
Pre-provisions OneDrive accounts for all licensed users in Commercial, GCC High, or DoD tenants.

.NOTES
- Requires Microsoft.Graph PowerShell SDK
- Requires SharePoint Online Management Shell
- MSOnline fallback works only in Commercial tenants
#>

Write-Host "==============================="
Write-Host " OneDrive Pre-Provisioning Tool "
Write-Host "==============================="

# Prompt for tenant type
Write-Host "Choose your tenant environment:"
Write-Host "1 = Commercial (O365 Global)"
Write-Host "2 = GCC / GCC High (USGov)"
Write-Host "3 = DoD / ITAR"
$tenantChoice = Read-Host "Enter 1, 2, or 3"

switch ($tenantChoice) {
    "1" {
        $graphEnv = "Global"
        $spoUrl = Read-Host "Enter SPO Admin URL (e.g. https://yourtenant-admin.sharepoint.com)"
        $spoParams = @{ Url = $spoUrl }
    }
    "2" {
        $graphEnv = "USGov"
        $spoUrl = Read-Host "Enter SPO Admin URL (e.g. https://yourtenant-admin.sharepoint.us)"
        $spoParams = @{ Url = "$spoUrl -Environment ITAR" }
    }
    "3" {
        $graphEnv = "USGovDoD"
        $spoUrl = Read-Host "Enter SPO Admin URL (e.g. https://yourtenant-admin.sharepoint-mil.us)"
        $spoParams = @{ Url = $spoUrl; Region = "ITAR" }
    }
    default {
        Write-Error "Invalid choice. Exiting."
        exit
    }
}

# Connect to SharePoint Online Admin service
Write-Host "🔗 Connecting to SharePoint Online..."
Connect-SPOService @spoParams

# Attempt Graph connection first
try {
    Import-Module Microsoft.Graph.Users -ErrorAction Stop
    Write-Host "🔗 Connecting to Microsoft Graph ($graphEnv)..."
    Connect-MgGraph -Scopes "User.Read.All" -Environment $graphEnv
    Write-Host "✅ Connected to Microsoft Graph"

    # Get all licensed users
    $licensedUsers = Get-MgUser -All -Property "userPrincipalName,assignedLicenses" | Where-Object {
        $_.AssignedLicenses -ne $null -and $_.AssignedLicenses.Count -gt 0
    }

    Write-Host "Found $($licensedUsers.Count) licensed users. Checking OneDrive sites..."

    foreach ($user in $licensedUsers) {
        $userEmail = $user.UserPrincipalName

        # Check if a OneDrive already exists for this user
        $existing = Get-SPOSite -IncludePersonalSite $true -Limit all | Where-Object { $_.Owner -eq $userEmail }

        if ($existing) {
            Write-Host "✔ $userEmail already has OneDrive. Skipping."
        }
        else {
            Write-Host "➡ Provisioning OneDrive for: $userEmail"
            Request-SPOPersonalSite -UserEmails $userEmail
        }
    }
}
catch {
    Write-Warning "Microsoft Graph connection failed or not installed."

    if ($tenantChoice -eq "1") {
        Write-Host "Falling back to MSOnline module (Commercial only)..."
        Import-Module MSOnline
        Connect-MsolService

        $licensedUsers = Get-MsolUser -All | Where-Object { $_.isLicensed -eq $true }

        foreach ($user in $licensedUsers) {
            $userEmail = $user.UserPrincipalName

            $existing = Get-SPOSite -IncludePersonalSite $true -Limit all | Where-Object { $_.Owner -eq $userEmail }

            if ($existing) {
                Write-Host "✔ $userEmail already has OneDrive. Skipping."
            }
            else {
                Write-Host "➡ Provisioning OneDrive for: $userEmail"
                Request-SPOPersonalSite -UserEmails $userEmail
            }
        }
    }
    else {
        Write-Error "❌ You must use Graph in GCC High/DoD tenants. MSOnline fallback not supported outside Commercial."
        exit
    }
}

Write-Host "======================================================"
Write-Host " ✅ OneDrive pre-provisioning completed successfully."
Write-Host "======================================================"

# Disconnect services
Disconnect-SPOService
Disconnect-MgGraph
