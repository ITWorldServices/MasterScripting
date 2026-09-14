# Import required modules
Import-Module ExchangeOnlineManagement
Import-Module ImportExcel
Import-Module Microsoft.Graph.Authentication
Import-Module Microsoft.Graph.Sites
Import-Module Microsoft.Graph.Users

# Connect to Exchange Online and Microsoft Graph
Connect-ExchangeOnline
Connect-MgGraph -Scopes "User.Read.All", "Sites.Read.All"

# Set the path for the output XLSX file
$exportPath = "C:\temp\MailboxPropertiesStats.xlsx"

# Define the properties you want to export in the desired order
$propertiesToExport = @(
    'DisplayName',
    'UserPrincipalName',
    'PrimarySmtpAddress',
    'RecipientTypeDetails',
    'AccountDisabled',
    'HiddenFromAddressListsEnabled',
    'DeliverToMailboxAndForward',
    'ForwardingAddress',
    'ForwardingSmtpAddress',
    'ItemCount',
    'TotalItemSize',
    'ArchiveStatus',
    'ArchiveTotalItemSize',
    'ArchiveItemCount'
    # Add more properties as needed
)

# Get all mailboxes directly from Exchange Online
Write-Host "Retrieving all mailboxes from Exchange Online..." -ForegroundColor Cyan
$mailboxList = Get-Mailbox -ResultSize Unlimited

# Create arrays to store data
$mailboxDataArray = @()
$notFoundMailboxes = @()

# Process each mailbox
$totalMailboxes = $mailboxList.Count
Write-Host "Processing $totalMailboxes mailboxes..." -ForegroundColor Cyan
$i = 0

foreach ($mailbox in $mailboxList) {
    $i++
    Write-Progress -Activity "Processing mailboxes" -Status "Processing $i of $totalMailboxes" -PercentComplete (($i / $totalMailboxes) * 100)
    
    try {
        $mailboxStats = Get-EXOMailboxStatistics -Identity $mailbox.UserPrincipalName -ErrorAction SilentlyContinue
        $archiveStats = $null
        if ($mailbox.ArchiveStatus -eq "Active") {
            $archiveStats = Get-EXOMailboxStatistics -Identity $mailbox.UserPrincipalName -Archive -ErrorAction SilentlyContinue
        }

        $combinedData = [PSCustomObject]@{
            DisplayName = $mailbox.DisplayName
            UserPrincipalName = $mailbox.UserPrincipalName
            PrimarySmtpAddress = $mailbox.PrimarySmtpAddress
            RecipientTypeDetails = $mailbox.RecipientTypeDetails
            AccountDisabled = $mailbox.AccountDisabled
            HiddenFromAddressListsEnabled = $mailbox.HiddenFromAddressListsEnabled
            DeliverToMailboxAndForward = $mailbox.DeliverToMailboxAndForward
            ForwardingAddress = $mailbox.ForwardingAddress
            ForwardingSmtpAddress = $mailbox.ForwardingSmtpAddress
            ItemCount = $mailboxStats.ItemCount
            TotalItemSize = $mailboxStats.TotalItemSize
            ArchiveStatus = $mailbox.ArchiveStatus
            ArchiveTotalItemSize = if ($archiveStats) { $archiveStats.TotalItemSize } else { "N/A" }
            ArchiveItemCount = if ($archiveStats) { $archiveStats.ItemCount } else { "N/A" }
        }

        $mailboxDataArray += $combinedData
    } catch {
        Write-Warning "Error processing mailbox: $($mailbox.UserPrincipalName). Error: $_"
        $notFoundMailboxes += $mailbox.UserPrincipalName
    }
}

# Export combined mailbox data to XLSX
Write-Host "Exporting data to Excel..." -ForegroundColor Cyan
$mailboxDataArray | Export-Excel -Path $exportPath -WorksheetName "Mailbox Data" -AutoSize -TableName "MailboxData" -TableStyle Medium1 -FreezeTopRow -BoldTopRow

# Apply formatting to the "MailboxData" worksheet
$excelPackage = Open-ExcelPackage -Path $exportPath
$worksheet = $excelPackage.Workbook.Worksheets["Mailbox Data"]

if ($worksheet) {
    # Set header style
    $worksheet.Cells["A1:N1"].Style.Fill.PatternType = [OfficeOpenXml.Style.ExcelFillStyle]::Solid
    $worksheet.Cells["A1:N1"].Style.Fill.BackgroundColor.SetColor([System.Drawing.Color]::SteelBlue)
    $worksheet.Cells["A1:N1"].Style.Font.Color.SetColor([System.Drawing.Color]::White)
    $worksheet.Cells["A1:N1"].Style.Font.Bold = $true

    # Auto-fit columns
    $worksheet.Cells.AutoFitColumns()

    Close-ExcelPackage $excelPackage
} else {
    Write-Warning "The 'MailboxData' worksheet was not created properly."
}

# Rest of the OneDrive export script remains unchanged...
