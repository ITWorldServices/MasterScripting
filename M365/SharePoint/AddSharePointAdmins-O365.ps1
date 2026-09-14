# Connect to SharePoint Online
Connect-SPOService -Url "https://icsholdingllc-admin.sharepoint.com"

# Specify the user you want to add as an administrator
$adminUser = "impactadmin@brifutelectric.com"

# Get a list of all non-personal site collections in the tenant
$sites = Get-SPOSite -Limit All | Where-Object { $_.Template -ne "SPSPERS" }

# Loop through each site and add the user as an administrator
foreach ($site in $sites) {
    Set-SPOSite -Identity $site.Url -Owner $adminUser
    Write-Host "User $adminUser added as an administrator to $($site.Url)"
}

Write-Host "Administrator added to all non-personal SharePoint Online sites."
