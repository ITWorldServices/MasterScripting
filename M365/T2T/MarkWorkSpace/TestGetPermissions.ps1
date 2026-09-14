<#
.SYNOPSIS
    Export SharePoint Online document library permissions for a specified site, excluding built-in libraries, and display results on screen.

.DESCRIPTION
    This script connects to a SharePoint Online site using PnP PowerShell, enumerates all document libraries except those specifed in the exclusion list,
    retrieves each library's user and group permissions, and outputs the results in both table format (on-screen) and a CSV file.
    The final export filename is prepended by the tenant name and site name, and appended with a timestamp for clear traceability.
    Common built-in libraries like 'Style Library', 'Form Templates', 'Site Assets', and 'Site Pages' are excluded, but you can modify the exclusion list as required.
    Designed for administrators auditing SharePoint permissions or reporting access for compliance and review.
#>

# Install PnP.PowerShell module if not installed:
# Install-Module PnP.PowerShell -Scope CurrentUser

$siteUrl    = "https://blackhawkcenter.sharepoint.com/teams/GoodwinAustinTeam"
$tenantName = "blackhawkcenter" # Prepend this to the file

# Extract site name from the URL
$siteName = ($siteUrl.Split('/')[-1])

# Get timestamp in YYYYMMDD_HHmm format
$timestamp = Get-Date -Format "yyyyMMdd_HHmm"

# Output filename (no library name included)
$outputPath = "${tenantName}_${siteName}_permissions_${timestamp}.csv"

Connect-PnPOnline -Url $siteUrl -ClientId "5e940a58-2d36-4e9e-9d6f-47009e181aea" -Interactive

# Exclude specified libraries
$ExcludedLibraries = @("Style Library", "Form Templates", "Site Assets", "Site Pages")
$lists = Get-PnPList | Where-Object {
    $_.BaseType -eq "DocumentLibrary" -and
    $_.Hidden -eq $false -and
    $_.Title -notin $ExcludedLibraries
}

$results = @()

foreach ($list in $lists) {
    $listObj = Get-PnPList -Identity $list.Title
    $roleAssignments = Get-PnPProperty -ClientObject $listObj -Property RoleAssignments

    foreach ($roleAssignment in $roleAssignments) {
        $nullSafeMember = Get-PnPProperty -ClientObject $roleAssignment -Property Member
        $nullSafeRoles  = Get-PnPProperty -ClientObject $roleAssignment -Property RoleDefinitionBindings

        $roles = $roleAssignment.RoleDefinitionBindings | Where-Object { $_ } | Select-Object -ExpandProperty Name
        $rolesString = ($roles -is [array]) ? ($roles -join ", ") : $roles

        if ($roleAssignment.Member -ne $null) {
            $principalType = $roleAssignment.Member.PrincipalType.ToString()
            $displayName   = $roleAssignment.Member.Title
        }
        else {
            $principalType = "Unknown"
            $displayName   = "Unknown"
        }

        $results += [PSCustomObject]@{
            Library        = $list.Title
            PrincipalType  = $principalType
            DisplayName    = $displayName
            Roles          = $rolesString
        }
    }
}

# Display results on the screen in a table format
$results | Format-Table -AutoSize

# Optionally, for interactive viewing:
# $results | Out-GridView

$results | Export-Csv -Path $outputPath -NoTypeInformation

Write-Host "Library permissions exported to $outputPath"
