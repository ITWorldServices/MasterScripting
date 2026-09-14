<#
.SYNOPSIS
Exports folders from a SharePoint Online document library up to a user-specified depth.

.DESCRIPTION
This PowerShell script connects to a SharePoint Online site using CSOM and exports folder information from any specified document library. 
The user is prompted for the site URL, document library name, and the maximum folder depth to export (from 1 to 5 levels, or 'all' for unlimited depth).
The script retrieves folder names, relative paths, URLs, and their hierarchical level, and exports this data to a CSV file.
The output file is named with the tenant name, library name, and a timestamp for easy identification.
This script is suitable for managing Microsoft tenants and supports flexible folder level selection[1].
#>

# Clear Screen
Clear

<# Load CSOM Assemblies
  Download the SharePoint Online Client Components SDK
       Use the official Microsoft download link:
       https://download.microsoft.com/download/B/3/D/B3DA6839-B852-41B3-A9DF-0AFA926242F2/sharepointclientcomponents_16-6906-1200_x64-en-us.msi
  Run the downloaded .msi installer on your production server.
#>
Add-Type -Path "C:\Program Files\Common Files\microsoft shared\Web Server Extensions\16\ISAPI\Microsoft.SharePoint.Client.dll"
Add-Type -Path "C:\Program Files\Common Files\microsoft shared\Web Server Extensions\16\ISAPI\Microsoft.SharePoint.Client.Runtime.dll"

# Prompt for Site URL and clean it
$SiteUrl = Read-Host "Enter your SharePoint Site URL (e.g., https://tenant.sharepoint.com/sites/yoursite)"
$SiteUrl = $SiteUrl -replace '[^\x00-\x7F]',''

# Validate Site URL format
if (-not [System.Uri]::IsWellFormedUriString($SiteUrl, [System.UriKind]::Absolute)) {
    throw "Invalid Site URL: '$SiteUrl'"
}

# Prompt for Document Library name
$LibraryName = Read-Host "Enter the Document Library name (case-sensitive)"

# Prompt for folder depth (1-5 or 'all')
$LevelInput = Read-Host "How many folder levels to export? (Enter 1-5 or 'all' for unlimited)"
if ($LevelInput -eq 'all') {
    $MaxDepth = [int]::MaxValue
} elseif ($LevelInput -match '^[1-5]$') {
    $MaxDepth = [int]$LevelInput
} else {
    throw "Invalid input for folder levels. Please enter 1, 2, 3, 4, 5, or 'all'."
}

# Extract tenant name and create timestamp
$tenantName = ([System.Uri]$SiteUrl).Host.Split('.')[0]
$timestamp = Get-Date -Format "yyyyMMdd_HHmmss"

# Securely prompt for credentials
$Cred = Get-Credential

# Create context
$Ctx = New-Object Microsoft.SharePoint.Client.ClientContext($SiteUrl)
$Ctx.Credentials = New-Object Microsoft.SharePoint.Client.SharePointOnlineCredentials($Cred.UserName, $Cred.Password)

# Define recursive function with proper parameters
function Get-FoldersRecursive {
    param (
        [Microsoft.SharePoint.Client.Folder]$ParentFolder,
        [Microsoft.SharePoint.Client.ClientContext]$Ctx,
        [string]$ParentPath = "",
        [int]$CurrentDepth = 1,
        [int]$MaxDepth
    )
    $folders = $ParentFolder.Folders
    $Ctx.Load($folders)
    $Ctx.ExecuteQuery()

    $results = @()
    foreach ($folder in $folders) {
        if ($folder.Name -ne "Forms" -and !$folder.Name.StartsWith("_")) {
            $fullPath = if ($ParentPath) { "$ParentPath/$($folder.Name)" } else { $folder.Name }
            $results += [PSCustomObject]@{
                Name = $folder.Name
                Path = $fullPath
                ServerRelativeUrl = $folder.ServerRelativeUrl
                Level = $CurrentDepth
            }
            if ($CurrentDepth -lt $MaxDepth) {
                $results += Get-FoldersRecursive -ParentFolder $folder -Ctx $Ctx -ParentPath $fullPath -CurrentDepth ($CurrentDepth + 1) -MaxDepth $MaxDepth
            }
        }
    }
    return $results
}

try {
    # Get document library
    $List = $Ctx.Web.Lists.GetByTitle($LibraryName)
    $Ctx.Load($List)
    $Ctx.Load($List.RootFolder)
    $Ctx.ExecuteQuery()

    # Get folders up to specified depth
    $RootFolder = $List.RootFolder
    $Ctx.Load($RootFolder)
    $Ctx.ExecuteQuery()
    
    # Call recursive function with all required parameters
    $Results = Get-FoldersRecursive -ParentFolder $RootFolder -Ctx $Ctx -MaxDepth $MaxDepth

    # Create filename with tenant, library, and timestamp
    $safeLibraryName = $LibraryName -replace '[\\\/:*?"<>| ]','_'
    $csvPath = "C:\temp\${tenantName}_${safeLibraryName}_Folders_$timestamp.csv"
    
    # Export to CSV
    $Results | Export-Csv -Path $csvPath -NoTypeInformation
    $Results | Format-Table
    Write-Host "Successfully exported $($Results.Count) folders (up to $LevelInput levels) to:"
    Write-Host $csvPath -ForegroundColor Cyan
}
catch {
    Write-Host "Error: $($_.Exception.Message)" -ForegroundColor Red
    Write-Host "Detailed error: $($_.ScriptStackTrace)" -ForegroundColor Yellow
}
