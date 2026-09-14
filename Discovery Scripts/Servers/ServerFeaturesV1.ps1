# Check if ImportExcel module is installed
$importExcelModuleInstalled = Get-Module -Name ImportExcel -ListAvailable

if (!$importExcelModuleInstalled) {
    Write-Host "ImportExcel module is not installed. Installing..."
    Install-Module -Name ImportExcel -Scope CurrentUser -Force
} else {
    Write-Host "ImportExcel module is already installed."
}

# Check if Active Directory module is installed
$adModuleInstalled = Get-Module -Name ActiveDirectory -ListAvailable

if (!$adModuleInstalled) {
    Write-Host "Active Directory module is not installed. Installing..."
    Install-WindowsFeature RSAT-AD-PowerShell
} else {
    Write-Host "Active Directory module is already installed."
}

# Import the required modules
Import-Module ImportExcel

# Specify the path to the CSV file containing server names
$csvPath = ".\serverlist.csv"

# Import the server names from the CSV file
$servers = Import-Csv -Path $csvPath | Select-Object -ExpandProperty ServerName

# Initialize an array to store the server feature information
$serverFeatureInfo = @()

# Loop through each server
foreach ($server in $servers) {
    try {
        # Print message to console
        Write-Host "Discovering Features for server: $server..."
        
        # Establish a remote session to the server
        $session = New-PSSession -ComputerName $server

        # Retrieve the installed server features
        $features = Invoke-Command -Session $session -ScriptBlock {
            Get-WindowsFeature | Where-Object { $_.Installed -eq "True" -and $_.FeatureType -eq "Feature" } | Select-Object -ExpandProperty Name
        }

        # Create a hashtable to store server name and features
        $featureInfo = [ordered]@{
            "Server" = $server
        }

        # Add each feature to the hashtable
        $index = 0
        foreach ($feature in $features) {
            $columnName = "Feature$index"
            $featureInfo[$columnName] = $feature
            $index++
        }

        # Add the feature information to the array
        $serverFeatureInfo += New-Object -TypeName PSObject -Property $featureInfo

        # Close the remote session
        Remove-PSSession -Session $session
    }
    catch {
        # Handle any errors that occur while connecting to a server
        Write-Host "Failed to connect to server: $server"
    }
}

# Export the server feature information to a formatted XLSX file with Blue Table Style Medium 2 format
$outputPath = "C:\Temp\ServerFeatures.xlsx"
$serverFeatureInfo | Export-Excel -Path $outputPath -AutoSize -AutoFilter -TableName "ServerFeatures" -TableStyle Medium2

# Display a success message
Write-Host "Server feature information exported to: $outputPath"
