# THIS SCRIPT EXPORTS MAILBOX PROPERTIES & STATS USING A CSV FILE AS THE USER LIST TO A xlsx FILE. 

# Ensure you have the Exchange Online PowerShell module installed
# If not installed, run this command:
# Install-Module -Name ExchangeOnlineManagement

# Ensure you have the ImportExcel module installed
# If not installed, run this command:
# Install-Module -Name ImportExcel -Force

# Import required modules
Import-Module ExchangeOnlineManagement
Import-Module ImportExcel

# Connect to Exchange Online
Connect-ExchangeOnline

#############################################################################
# Set the path for the input CSV file                                       #
# NOTE: CSV FILE MUST CONTAIN THREE HEADERS; UserPrincipalName, Mail, Email #
#############################################################################
$inputPath = "c:\temp\GEM-Source-UPN.csv"

#########################################
# Set the path for the output XLSX file #
#########################################
$exportPath = "C:\temp\NEW-GEM-MailboxPropertiesAndStats.xlsx"

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

# Import the CSV file
$mailboxList = Import-Csv -Path $inputPath

# Create arrays to store data
$mailboxDataArray = @()
$notFoundMailboxes = @()

# Loop through each mailbox in the CSV and collect properties and statistics
foreach ($mailboxEntry in $mailboxList) {
    $mailbox = Get-Mailbox -Identity $mailboxEntry.UserPrincipalName -ErrorAction SilentlyContinue
    if ($mailbox) {
        $mailboxProperties = $mailbox | Select-Object $propertiesToExport
        $mailboxStats = Get-EXOMailboxStatistics -Identity $mailboxEntry.UserPrincipalName -ErrorAction SilentlyContinue
        $archiveStats = $null
        if ($mailbox.ArchiveStatus -eq "Active") {
            $archiveStats = Get-EXOMailboxStatistics -Identity $mailboxEntry.UserPrincipalName -Archive -ErrorAction SilentlyContinue
        }

        $combinedData = [PSCustomObject]@{
            DisplayName = $mailboxProperties.DisplayName
            UserPrincipalName = $mailboxProperties.UserPrincipalName
            PrimarySmtpAddress = $mailboxProperties.PrimarySmtpAddress
            RecipientTypeDetails = $mailboxProperties.RecipientTypeDetails
            AccountDisabled = $mailboxProperties.AccountDisabled
            HiddenFromAddressListsEnabled = $mailboxProperties.HiddenFromAddressListsEnabled
            DeliverToMailboxAndForward = $mailboxProperties.DeliverToMailboxAndForward
            ForwardingAddress = $mailboxProperties.ForwardingAddress
            ForwardingSmtpAddress = $mailboxProperties.ForwardingSmtpAddress
            ItemCount = $mailboxStats.ItemCount
            TotalItemSize = $mailboxStats.TotalItemSize
            ArchiveStatus = $mailbox.ArchiveStatus
            ArchiveTotalItemSize = if ($archiveStats) { $archiveStats.TotalItemSize } else { "N/A" }
            ArchiveItemCount = if ($archiveStats) { $archiveStats.ItemCount } else { "N/A" }
        }

        $mailboxDataArray += $combinedData
    } else {
        Write-Warning "Mailbox not found: $($mailboxEntry.UserPrincipalName)"
        $notFoundMailboxes += $mailboxEntry.UserPrincipalName
    }
}


#########################
# EXPORT TO SPREADSHEET #                                                                                                              #
#########################
# Export combined mailbox data to XLSX
$mailboxDataArray | Export-Excel -Path $exportPath -WorksheetName "Mailbox Data" -AutoSize -TableName "MailboxData" -TableStyle Medium1 -FreezeTopRow -BoldTopRow

# Apply formatting to the "MailboxeData" worksheet
$excelPackage = Open-ExcelPackage -Path $exportPath
$worksheet = $excelPackage.Workbook.Worksheets["Mailbox Data"]

if ($worksheet) {
    # Set header style
    $worksheet.Cells["A1:M1"].Style.Fill.PatternType = [OfficeOpenXml.Style.ExcelFillStyle]::Solid
    $worksheet.Cells["A1:M1"].Style.Fill.BackgroundColor.SetColor([System.Drawing.Color]::SteelBlue)
    $worksheet.Cells["A1:M1"].Style.Font.Color.SetColor([System.Drawing.Color]::White)
    $worksheet.Cells["A1:M1"].Style.Font.Bold = $true

    # Auto-fit columns
    $worksheet.Cells.AutoFitColumns()

    Close-ExcelPackage $excelPackage -Show
} else {
    Write-Warning "The 'MailboxData' worksheet was not created properly. This might happen if There are no mailboxes or something went very wrong."
}

# Prepare data for mailboxes not found
$notFoundCount = $notFoundMailboxes.Count
$notFoundData = @(
    [PSCustomObject]@{
        "Number of mailboxes not found" = $notFoundCount
    }
)

# Add each not found mailbox to the data
foreach ($mailbox in $notFoundMailboxes) {
    $notFoundData += [PSCustomObject]@{
        "Mailboxes not found" = $mailbox
    }
}

# Export not found mailboxes data to a new worksheet
$notFoundData | Export-Excel -Path $exportPath -WorksheetName "Not Found Mailboxes" -AutoSize -TableName "NotFoundMailboxes" -TableStyle Medium1 -FreezeTopRow -BoldTopRow -Append
$notFoundMailboxes | Export-Excel -Path $exportPath -WorksheetName "Not Found Mailboxes" -AutoSize -TableName "NotFoundMailboxes" -TableStyle Medium1 -FreezeTopRow -BoldTopRow -Append

# Apply formatting to the "Not Found Mailboxes" worksheet
$excelPackage = Open-ExcelPackage -Path $exportPath
$worksheet = $excelPackage.Workbook.Worksheets["Not Found Mailboxes"]

if ($worksheet) {
    # Set header style
    $worksheet.Cells["A1:A1"].Style.Fill.PatternType = [OfficeOpenXml.Style.ExcelFillStyle]::Solid
    $worksheet.Cells["A1:A1"].Style.Fill.BackgroundColor.SetColor([System.Drawing.Color]::ste)
    $worksheet.Cells["A1:A1"].Style.Font.Color.SetColor([System.Drawing.Color]::White)
    $worksheet.Cells["A1:A1"].Style.Font.Bold = $true

    # Auto-fit columns
    $worksheet.Cells.AutoFitColumns()

    Close-ExcelPackage $excelPackage -Show
} else {
    Write-Warning "The 'Not Found Mailboxes' worksheet was not created. This might happen if no mailboxes were missing."
}
