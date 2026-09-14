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

# Import the required module
Import-Module ImportExcel

# Specify the path to the CSV file containing server names
$csvPath = ".\serverlist.csv"

# Import the server names from the CSV file
$servers = Import-Csv -Path $csvPath | Select-Object -ExpandProperty ServerName

# Initialize an array to store the installed application information
$serverAppInfo = @()

# Loop through each server
foreach ($server in $servers) {
    try {
        # Print message to console
        Write-Host "Discovering applications for server: $server..."

        # Establish a remote session to the server
        $session = New-PSSession -ComputerName $server

        # Retrieve the installed applications
        $applications = Invoke-Command -Session $session -ScriptBlock {
            Get-WmiObject -Class Win32_Product | Select-Object -Property Name
        }

        # Create a hashtable to store server name and applications
        $appInfo = [ordered]@{
            "Server" = $server
        }

        # Add each application to the hashtable
        $index = 0
        foreach ($application in $applications) {
            $columnName = "Application$index"
            $appInfo[$columnName] = $application.Name
            $index++
        }

        # Add the application information to the array
        $serverAppInfo += New-Object -TypeName PSObject -Property $appInfo

        # Close the remote session
        Remove-PSSession -Session $session
    }
    catch {
        # Handle any errors that occur while connecting to a server
        Write-Host "Failed to connect to server: $server"
    }
}

# Export the installed application information to a formatted XLSX file with Blue Table Style Medium 2 format
$outputPath = "C:\Temp\ServerApplications.xlsx"
$serverAppInfo | Export-Excel -Path $outputPath -AutoSize -AutoFilter -TableName "ServerApplications" -TableStyle Medium2

# Display a success message
Write-Host "Installed application information exported to: $outputPath"
