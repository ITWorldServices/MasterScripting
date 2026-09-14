# Import required modules
Import-Module ExchangeOnlineManagement
Import-Module ImportExcel
Import-Module Microsoft.Graph.Authentication
Import-Module Microsoft.Graph.Sites
Import-Module Microsoft.Graph.Users

# Connect to Exchange Online and Microsoft Graph
Connect-ExchangeOnline
Connect-MgGraph -Scopes "User.Read.All", "Sites.Read.All"

#####################################################################################################################
# Set the path for the input CSV file                                                                               #                             
# NOTE: CSV FILE MUST CONTAIN THREE HEADERS; UserPrincipalName, Mail, Email (MAKE SURE TO USE "_" FOR ALL SEPARATOR #
#####################################################################################################################
$inputPath = "c:\temp\i3d-Source-UPN.csv"

#########################################
# Set the path for the output XLSX file #
#########################################
$exportPath = "C:\temp\NEW-I3D-Mailbox-Properties-Stats-OneDriveUsage.xlsx"

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

# Export combined mailbox data to XLSX
$mailboxDataArray | Export-Excel -Path $exportPath -WorksheetName "Mailbox Data" -AutoSize -TableName "MailboxData" -TableStyle Medium1 -FreezeTopRow -BoldTopRow

# Apply formatting to the "MailboxData" worksheet
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

    Close-ExcelPackage $excelPackage
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
    $worksheet.Cells["A1:A1"].Style.Fill.BackgroundColor.SetColor([System.Drawing.Color]::SteelBlue)
    $worksheet.Cells["A1:A1"].Style.Font.Color.SetColor([System.Drawing.Color]::White)
    $worksheet.Cells["A1:A1"].Style.Font.Bold = $true

    # Auto-fit columns
    $worksheet.Cells.AutoFitColumns()

    Close-ExcelPackage $excelPackage
} else {
    Write-Warning "The 'Not Found Mailboxes' worksheet was not created. This might happen if no mailboxes were missing."
}

###########################
# OneDrive Export Section #
###########################

# Get all users
$users = Get-MgUser -All

# Initialize an array to store the results
$results = @()

foreach ($user in $users) {
    try {
        # Get OneDrive site for the user
        $oneDriveSite = Get-MgUserDrive -UserId $user.Id -ErrorAction Stop

        # Get OneDrive usage details
        $driveUsage = Get-MgDrive -DriveId $oneDriveSite.Id

        $result = [PSCustomObject]@{
            User = $user.UserPrincipalName
            TotalSizeMB = [math]::Round($driveUsage.Quota.Used / 1MB, 2)
            ProvisioningStatus = "Provisioned"
            LastAccessTime = $driveUsage.LastModifiedDateTime
        }
    }
    catch {
        # If OneDrive is not provisioned, add user with empty data
        $result = [PSCustomObject]@{
            User = $user.UserPrincipalName
            TotalSizeGB = 0
            ItemCount = 0
            ProvisioningStatus = "Not Provisioned"
            LastAccessTime = $null
        }
    }

    $results += $result
}

# Export OneDrive data to a new worksheet in the existing Excel file
$results | Export-Excel -Path $exportPath -WorksheetName "OneDrive Usage" -AutoSize -TableName "OneDriveUsage" -TableStyle Medium1 -FreezeTopRow -BoldTopRow -Append

# Apply formatting to the "OneDrive Usage" worksheet
$excelPackage = Open-ExcelPackage -Path $exportPath
$worksheet = $excelPackage.Workbook.Worksheets["OneDrive Usage"]

if ($worksheet) {
    # Set header style
    $worksheet.Cells["A1:D1"].Style.Fill.PatternType = [OfficeOpenXml.Style.ExcelFillStyle]::Solid
    $worksheet.Cells["A1:D1"].Style.Fill.BackgroundColor.SetColor([System.Drawing.Color]::SteelBlue)
    $worksheet.Cells["A1:D1"].Style.Font.Color.SetColor([System.Drawing.Color]::White)
    $worksheet.Cells["A1:D1"].Style.Font.Bold = $true

    # Auto-fit columns
    $worksheet.Cells.AutoFitColumns()

    Close-ExcelPackage $excelPackage -Show
} else {
    Write-Warning "The 'OneDrive Usage' worksheet was not created properly."
}

# Display summary
$totalUsers = $results.Count
$provisionedUsers = ($results | Where-Object { $_.ProvisioningStatus -eq "Provisioned" }).Count
$notProvisionedUsers = $totalUsers - $provisionedUsers

Write-Host "`nSummary:"
Write-Host "Total Users: $totalUsers"
Write-Host "Provisioned OneDrives: $provisionedUsers"
Write-Host "Not Provisioned OneDrives: $notProvisionedUsers"
Disconnect-ExchangeOnline
Disconnect-MgGraph