#Script to test the bandwidth at a site to two different servers, average them and insert the results in a Bandwidth tab within a spreadsheet. 

$speedtestDownloaded = $false
$importExcelModuleLoaded = $false

# Check and create temp folder
$folderPath = "C:\temp"
if (-not (Test-Path $folderPath)) {
    New-Item -Path $folderPath -ItemType Directory
}

# Only download and extract speedtest.exe if it doesn't exist
if (-not (Test-Path "$folderPath\speedtest.exe")) {
    $zipPath = "$folderPath\ookla-speedtest-win64.zip"
    $downloadUrl = "https://install.speedtest.net/app/cli/ookla-speedtest-1.2.0-win64.zip"

    Invoke-WebRequest -Uri $downloadUrl -OutFile $zipPath
    Expand-Archive -Path $zipPath -DestinationPath $folderPath -Force
	$speedtestDownloaded=$true
}

# Retrieve the list of servers
$serversOutput = & "C:\temp\speedtest.exe" -L
Write-Host "Scanning for available servers and selecting the top two entries. Standby..."

# Match and capture server IDs
$matches = [regex]::Matches($serversOutput, "\b(\d{4,5})\b")

# Extract the server IDs
$serverIDs = $matches | ForEach-Object { $_.Groups[1].Value } | Select-Object -First 2

# Initialize results array
$results = @()

# Run speed test for each of the top two server IDs
foreach ($id in $serverIDs) {
    Start-Sleep -Seconds 30
	Write-Host "Processing ID: $id Please Standby..."
    $testOutput = & "C:\temp\speedtest.exe" --server-id=$id --format=json --accept-license --accept-gdpr
    $testJson = $testOutput | ConvertFrom-Json

    $results += [PSCustomObject]@{
        "Test Label"  = "Test-$id"
        ExternalIP    = $testJson.interface.externalIp
        isp           = $testJson.isp
        InternalIP    = $testJson.interface.internalIp
        Mbpsdownloadspeed = [math]::Round($testJson.download.bandwidth / 1000000 * 8, 2)
        Mbpsuploadspeed   = [math]::Round($testJson.upload.bandwidth / 1000000 * 8, 2)
        packetloss    = [math]::Round($testJson.packetLoss)
        Jitter        = [math]::Round($testJson.ping.jitter)
        Latency       = [math]::Round($testJson.ping.latency)
        UsedServer    = $testJson.server.host
    }
}

# Calculate average and add to results
$results += [PSCustomObject]@{
    "Test Label"  = "Average"
    ExternalIP    = ""
    isp           = ""
    InternalIP    = ""
    Mbpsdownloadspeed = ($results.Mbpsdownloadspeed | Measure-Object -Average).Average
    Mbpsuploadspeed   = ($results.Mbpsuploadspeed | Measure-Object -Average).Average
    packetloss    = ($results.packetloss | Measure-Object -Average).Average
    Jitter        = ($results.Jitter | Measure-Object -Average).Average
    Latency       = ($results.Latency | Measure-Object -Average).Average
    UsedServer    = ""
}

#START THE EXPORT OF THE DATA BY MOVING IT TO A CSV THEN IMPORTING THAT CSV INTO AN EXCEL spreadsheet

# Output the results to a CSV file
$csvFilePath = "$folderPath\SpeedTestResults.csv"
$results | Export-Csv -Path $csvFilePath -NoTypeInformation -Delimiter ","

# Display the results in console
$results | Format-Table

# Define the path to the Excel file
$excelFilePath = "C:\temp\TPT - Exchange to O365 Mail Migration Planning Document.xlsx"

# Check if the ImportExcel module is installed
if (-not (Get-Module -ListAvailable -Name ImportExcel)) {
    # If not installed, install the module
    Install-Module -Name ImportExcel -Scope CurrentUser -Force -Confirm:$false
	$importExcelModuleLoaded = $true

    if (-not (Get-Module -ListAvailable -Name ImportExcel)) {
        Write-Error "Failed to install the ImportExcel module. Exiting script."
        return
    }
}

# Import the module
Import-Module ImportExcel

# Check if the Excel file exists
if (Test-Path $excelFilePath) {
    # Check if the "Bandwidth" tab exists in the Excel file
    $existingSheets = (Open-ExcelPackage -Path $excelFilePath).Workbook.Worksheets.Name
    if ('Bandwidth' -notin $existingSheets) {
        # Create the "Bandwidth" tab if it doesn't exist
        New-ExcelWorksheet -Path $excelFilePath -WorksheetName "Bandwidth"
    }
} else {
    Write-Error "Excel file not found at $excelFilePath. Exiting script."
    return
}

# Import the CSV data into the "Bandwidth" tab
Import-Csv -Path $csvFilePath | Export-Excel -Path $excelFilePath -WorksheetName "Bandwidth" -AutoSize

Write-Host "Data sent to Excel file successfully!"

#CLEANUP!!

# Cleanup
if ($importExcelModuleLoaded) {
    # Unload the ImportExcel module if it was loaded during the script execution
    Remove-Module ImportExcel
}

# Check if speedtest.exe was downloaded during this script's execution
if ($speedtestDownloaded) {
    # Remove speedtest.exe if it was downloaded during the script execution
    Remove-Item -Path "C:\temp\speedtest.exe" -Force
}

# Check if the CSV file exists and then remove it
if (Test-Path $csvFilePath) {
    Remove-Item -Path $csvFilePath -Force
}
