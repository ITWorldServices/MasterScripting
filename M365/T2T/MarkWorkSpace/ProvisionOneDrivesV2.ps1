# PowerShell Script: Provision OneDrive Sites for Single or Multiple Users
# Description: Checks if user’s OneDrive is already provisioned and supports single or CSV modes.

$cacheFile = "$env:TEMP\SPOAdminUrlCache.txt"

# Load cached URL
if (Test-Path $cacheFile) {
    $cachedUrl = Get-Content $cacheFile -ErrorAction SilentlyContinue
    $adminUrl = Read-Host "Enter your SharePoint Admin Center URL [Press Enter to reuse: $cachedUrl]"

    if ([string]::IsNullOrWhiteSpace($adminUrl)) {
        $adminUrl = $cachedUrl
        Write-Host "Reusing cached admin URL: $adminUrl" -ForegroundColor Yellow
    }
    else {
        Set-Content -Path $cacheFile -Value $adminUrl
    }
}
else {
    $adminUrl = Read-Host "Enter your SharePoint Admin Center URL (e.g., https://contoso-admin.sharepoint.com)"
    if (-not [string]::IsNullOrWhiteSpace($adminUrl)) {
        Set-Content -Path $cacheFile -Value $adminUrl
    }
}

if ([string]::IsNullOrWhiteSpace($adminUrl)) {
    Write-Host "No admin URL entered and no cached URL found. Exiting." -ForegroundColor Red
    exit
}

Connect-SPOService -Url $adminUrl

function Check-OneDriveExists {
    param([string]$userEmail)
    try {
        $existingSite = Get-SPOSite -IncludePersonalSite $true -Filter "Owner -eq '$userEmail'" -Limit All -ErrorAction SilentlyContinue
        if ($existingSite) {
            Write-Host ("OneDrive already exists for {0}: {1}" -f $userEmail, $existingSite.Url) -ForegroundColor Yellow
            return $true
        }
        else {
            return $false
        }
    }
    catch {
        Write-Host ("Error checking OneDrive for {0}: {1}" -f $userEmail, $_.Exception.Message) -ForegroundColor Red
        return $false
    }
}

$mode = Read-Host "Do you want to provision a single user or multiple users from a CSV? Enter 'Single' or 'CSV'"

if ($mode -eq 'Single') {
    do {
        $userEmail = Read-Host "Enter the user's email address"

        if ([string]::IsNullOrWhiteSpace($userEmail)) {
            Write-Host "Invalid email. Please try again." -ForegroundColor Red
            continue
        }

        Write-Host "Checking OneDrive status for: $userEmail" -ForegroundColor Cyan
        $exists = Check-OneDriveExists -userEmail $userEmail

        if ($exists) {
            Write-Host "Skipping $userEmail — OneDrive already provisioned." -ForegroundColor Yellow
        }
        else {
            Write-Host "Provisioning OneDrive for: $userEmail" -ForegroundColor Cyan
            try {
                Request-SPOPersonalSite -UserEmails $userEmail
                Write-Host "Successfully provisioned OneDrive for $userEmail." -ForegroundColor Green
            }
            catch {
                Write-Host "Failed to provision OneDrive for $userEmail. Error: $_" -ForegroundColor Red
            }
        }

        $continue = Read-Host "Do you want to provision another single user? (Y/N)"
    } while ($continue -eq 'Y')
}
elseif ($mode -eq 'CSV') {
    $csvPath = Read-Host "Enter the full path to your CSV file (e.g., C:\Temp\users.csv)"
    if (Test-Path $csvPath) {
        $csvData = Import-Csv -Path $csvPath
        foreach ($row in $csvData) {
            $userEmail = $row.Email
            if (-not $userEmail) { continue }

            Write-Host "Checking OneDrive status for: $userEmail" -ForegroundColor Yellow
            $exists = Check-OneDriveExists -userEmail $userEmail

            if ($exists) {
                Write-Host "Skipping $userEmail — OneDrive already provisioned." -ForegroundColor Yellow
            }
            else {
                Write-Host "Provisioning OneDrive for: $userEmail" -ForegroundColor Cyan
                try {
                    Request-SPOPersonalSite -UserEmails $userEmail
                    Write-Host "Successfully provisioned OneDrive for $userEmail." -ForegroundColor Green
                }
                catch {
                    Write-Host "Failed to provision OneDrive for $userEmail. Error: $_" -ForegroundColor Red
                }
            }
        }
        Write-Host "OneDrive pre-provisioning completed for all users." -ForegroundColor Cyan
    }
    else {
        Write-Host "CSV file not found. Please check the path." -ForegroundColor Red
    }
}
else {
    Write-Host "Invalid selection. Please run again and enter 'Single' or 'CSV'." -ForegroundColor Red
}

Disconnect-SPOService
