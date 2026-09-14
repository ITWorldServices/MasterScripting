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
Import-Module ActiveDirectory
Import-Module ImportExcel

# Get the current domain name
$domainName = (Get-ADDomain).DNSRoot

# Discover all Active Directory Domain Controllers
$domainControllers = Get-ADDomainController -Filter * -Server $domainName | Select-Object -ExpandProperty Name

# Initialize an array to store the Domain Controller information
$dcInfo = @()

# Loop through each Domain Controller
foreach ($dc in $domainControllers) {
    # Print message to console
        Write-Host "Discovering Domain Controller Information: $dc..."
    
    # Retrieve the DHCP installation status
    $dhcpInstalled = Get-Service -ComputerName $dc -Name "DHCPServer" -ErrorAction SilentlyContinue

    # Retrieve the Read-Only DC state as a boolean value
    $readOnlyDC = [bool](Get-ADDomainController -Identity $dc).ReadOnly

    # Create a hashtable to store Domain Controller information
    $dcData = [ordered]@{
        "FQDN"           = $dc
        "OperatingSystem"   = (Get-ADDomainController -Identity $dc).OperatingSystem
        "Status"         = $true
        "GlobalCatalog"   = (Get-ADDomainController -Identity $dc).IsGlobalCatalog
        "Site"           = (Get-ADDomainController -Identity $dc).Site
        "ReadOnlyDC"     = $readOnlyDC
        "DHCPInstalled" = $dhcpInstalled -ne $null
    }

    # Add the Domain Controller information to the array
    $dcInfo += New-Object -TypeName PSObject -Property $dcData
}

# Export the Domain Controller information to a formatted XLSX file with Blue Table Style Medium 2 format
$outputPath = "C:\Temp\DomainControllers.xlsx"
$dcInfo | Export-Excel -Path $outputPath -AutoSize -AutoFilter -TableName "DomainControllers" -TableStyle Medium2

# Display a success message
Write-Host "Domain Controller information exported to: $outputPath"
