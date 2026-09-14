## Script to Compile Output

# Specify the path of the directory that contains the workbooks
$workbookDirectory = "C:\Temp"

# Specify the names of the workbooks
$workbookList = @("ServerInfo.xlsx", "ServerApplications.xlsx", "DomainControllers.xlsx", "ServerFeatures.xlsx", "ServerRoles.xlsx") # Add or remove workbook names as needed

# Specify the output workbook name
$outputWorkbook = "ServerPlanningSheet.xlsx"

# Build the full path of the output workbook
$outputWorkbookPath = Join-Path -Path $workbookDirectory -ChildPath $outputWorkbook

# Iterate over each workbook in the list
foreach ($workbook in $workbookList) {
    # Build the full path of the workbook
    $workbookPath = Join-Path -Path $workbookDirectory -ChildPath $workbook

    # Import the data from the workbook
    $data = Import-Excel -Path $workbookPath

    # Define the worksheet name based on the workbook name
    $worksheetName = [System.IO.Path]::GetFileNameWithoutExtension($workbook)

    # Export the data to the output workbook (appending to existing data)
    $data | Export-Excel -Path $outputWorkbookPath -WorksheetName $worksheetName -AutoSize -AutoFilter -Append
}

# Display a success message
Write-Host "Data exported to: $outputWorkbookPath"
