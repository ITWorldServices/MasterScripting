# Set your default or last used site URL
$lastSiteUrl = "https://blackhawkcenter.sharepoint.com/teams/GoodwinSanAntonio"

# Prompt for the site URL, use previous value if blank
$inputSiteUrl = Read-Host "Enter the site URL [$lastSiteUrl is the default, press Enter to accept]"
if ([string]::IsNullOrWhiteSpace($inputSiteUrl)) {
    $siteUrl = $lastSiteUrl
} else {
    $siteUrl = $inputSiteUrl
    $lastSiteUrl = $inputSiteUrl  # Update for next use (persist externally if needed)
}

# Connect to SharePoint
Connect-PnPOnline -Url $siteUrl -ClientId "5e940a58-2d36-4e9e-9d6f-47009e181aea" -Interactive

# List existing document libraries
Write-Host "`nExisting Document Libraries in ${siteUrl}:" -ForegroundColor Cyan
Get-PnPList | Where-Object { $_.BaseType -eq "DocumentLibrary" -and $_.Hidden -eq $false } | Select-Object Title, Url | Format-Table -AutoSize

# Import folder names from CSV
$folders = Import-Csv -Path "C:\Temp\SANANTONIO-folders.csv"

# Create folders in specified library
foreach ($folder in $folders) {
    Add-PnPFolder -Name $folder.FolderName -Folder $folder.root
}
