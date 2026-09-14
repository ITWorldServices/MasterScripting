# Set your default or last-used site URL
$lastSiteUrl = "https://blackhawkcenter.sharepoint.com/teams/GoodwinDallas"

# Prompt for the site URL
$inputSiteUrl = Read-Host "Enter the site URL [$lastSiteUrl is the default, press Enter to accept]"
if ([string]::IsNullOrWhiteSpace($inputSiteUrl)) {
    $siteUrl = $lastSiteUrl
} else {
    $siteUrl = $inputSiteUrl
    $lastSiteUrl = $inputSiteUrl
}

# Connect to SharePoint
Connect-PnPOnline -Url $siteUrl -ClientId "5e940a58-2d36-4e9e-9d6f-47009e181aea" -Interactive

# List existing document libraries
Write-Host "`nExisting Document Libraries in ${siteUrl}:" -ForegroundColor Cyan
$libs = Get-PnPList | Where-Object { $_.BaseType -eq "DocumentLibrary" -and $_.Hidden -eq $false } | Select-Object Title, RootFolder

$libs | Format-Table Title, RootFolder -AutoSize

# Prompt for the location of the CSV file
$csvPath = Read-Host "Enter the full path to the CSV file of folder names"

# Import folder names from CSV
$folders = Import-Csv -Path $csvPath

# Display folder creation details for confirmation
Write-Host "`nFolders will be added to the libraries"
Write-Host "Source CSV: $csvPath"
Write-Host "Number of folders: " $folders.Count
$confirm = Read-Host "Proceed with folder creation? (Y/N)"

if ($confirm -eq 'Y') {
    foreach ($folder in $folders) {
        Add-PnPFolder -Name $folder.FolderName -Folder $folder.root
    }
    Write-Host "Folders added successfully."
} else {
    Write-Host "Operation cancelled by user."
}
