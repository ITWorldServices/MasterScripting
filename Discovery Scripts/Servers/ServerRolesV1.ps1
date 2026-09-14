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

# Initialize a hashtable to store the server role information
$serverRoleInfo = @{}

# Loop through each server
foreach ($server in $servers) {
    try {
         # Print message to console
        Write-Host "Discovering Features for server: $server..."
        
        # Establish a remote session to the server
        $session = New-PSSession -ComputerName $server

        # Retrieve the installed server roles
        $roles = Invoke-Command -Session $session -ScriptBlock {
            Get-WindowsFeature | Where-Object { $_.Installed -eq "True" } | Select-Object -ExpandProperty Name
        }

        # Add the server and roles to the hashtable
        $serverRoleInfo[$server] = $roles

        # Close the remote session
        Remove-PSSession -Session $session
    }
    catch {
        # Handle any errors that occur while connecting to a server
        Write-Host "Failed to connect to server: $server"
    }
}

# Create an array of objects to hold the formatted server role information
$exportData = @()
foreach ($server in $serverRoleInfo.Keys) {
    $roles = $serverRoleInfo[$server]
    $row = [PSCustomObject]@{
        "Server" = $server
    }
    foreach ($index in 0..($roles.Count - 1)) {
        $columnName = "Role$index"
        $row | Add-Member -MemberType NoteProperty -Name $columnName -Value $roles[$index]
    }
    $exportData += $row
}

# Export the server role information to a formatted XLSX file with Blue Table Style Medium 2 format
$outputPath = "C:\Temp\ServerRoles.xlsx"
$exportData | Export-Excel -Path $outputPath -AutoSize -AutoFilter -TableName "ServerRoles" -TableStyle Medium2

# Display a success message
Write-Host "Server role information exported to: $outputPath"
