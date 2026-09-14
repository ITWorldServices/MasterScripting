# Connect to SharePoint
Connect-PnPOnline -Url "https://blackhawkcenter.sharepoint.com/teams/GoodwinAustinTeam" -ClientId "5e940a58-2d36-4e9e-9d6f-47009e181aea" -Interactive

# Import folder names from CSV
$folders = Import-Csv -Path "C:\Temp\D-folders.csv"

foreach ($folder in $folders) {
    Add-PnPFolder -Name $folder.FolderName -Folder "d"
}
