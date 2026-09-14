# PreProvision OneDrive for Licensed Users (Commercial + GCC High/DoD)

# Prompt for Tenant Type
Write-Host "Choose your tenant environment:"
Write-Host "1 = Commercial (O365 Global)"
Write-Host "2 = GCC / GCC High (USGov)"
Write-Host "3 = DoD / ITAR"
$tenantChoice = Read-Host "Enter 1, 2, or 3"

switch ($tenantChoice) {
    "1" {
        $graphEnv = "Global"
        $spoUrl = Read-Host "Enter SPO Admin URL (e.g. https://yourtenant-admin.sharepoint.com)"
    }
    "2" {
        $graphEnv = "USGov"
        $spoUrl = Read-Host "Enter SPO Admin URL (e.g. https://yourtenant-admin.sharepoint.us)"
    }
    "3" {
        $graphEnv = "USGovDoD"
        $spoUrl = Read-Host "Enter SPO Admin URL (e.g. https://yourtenant-admin.sharepoint-mil.us)"
    }
    default {
        Write-Error "Invalid choice. Exiting."
        exit
    }
}

# Connect to SharePoint Online
Connect-SPOService -Url $spoUrl

# Attempt Microsoft Graph Connection (preferred)
try {
    Import-Module Microsoft.Graph.Users -ErrorAction Stop
    Connect-MgGraph -Environment $graphEnv -Scopes "User.Read.All"
    Write-Host "✅ Connected to Microsoft Graph ($graphEnv)"
    
    # Collect licensed users
    $licensedUsers = Get-MgUser -All -Property "userPrincipalName,assignedLicenses" | Where-Object {
        $_.AssignedLicenses -ne $null -and $_.AssignedLicenses.Count -gt 0
    }

    Write-Host "Found $($licensedUsers.Count) licensed users. Checking OneDrive provisioning status..."

    foreach ($user in $licensedUsers) {
        $userEmail = $user.UserPrincipalName

        # Check if OneDrive already exists
        try {
            Get-SPOSite -IncludePersonalSite $true -Filter "Owner -eq '$userEmail'" -ErrorAction Stop | Out-Null
            Write-Host "✔ $userEmail - OneDrive already exists, skipping..."
        }
        catch {
            Write-Host "➡ Provisioning OneDrive for: $userEmail"
            Request-SPOPersonalSite -UserEmails $userEmail
        }
    }

} catch {
    Write-Host "❌ Microsoft Graph not available. Falling back to MSOnline (commercial only)"
    if ($tenantChoice -ne "1") {
        Write-Error "MSOnline does not support GCC/DoD. You must use Graph in USGov/DoD tenants."
        exit
    }

    # Fallback for Commercial tenants only
    Import-Module MSOnline
    Connect-MsolService

    $licensedUsers = Get-MsolUser -All | Where-Object { $_.isLicensed -eq $true }

    foreach ($user in $licensedUsers) {
        $userEmail = $user.UserPrincipalName

        try {
            Get-SPOSite -IncludePersonalSite $true -Filter "Owner -eq '$userEmail'" -ErrorAction Stop | Out-Null
            Write-Host "✔ $userEmail - OneDrive already exists, skipping..."
        }
        catch {
            Write-Host "➡ Provisioning OneDrive for: $userEmail"
            Request-SPOPersonalSite -UserEmails $userEmail
        }
    }
}

Write-Host "✅ OneDrive preprovisioning completed for all applicable users."

Disconnect-SPOService
