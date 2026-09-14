# Set parameters
param (
    [Parameter(Mandatory = $false)]
    [switch]$Reinstall
)

# Uninstall registry path
$registryPath = "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\"

# Initialize variables
$uninstallString = $null

# Get all subkeys in the uninstall registry path
$uninstallkeys = Get-ChildItem -Path $registryPath

# Loop through and find the uninstall string for INC
foreach ($key in $uninstallkeys) {
    $displayName = (Get-ItemProperty -Path $key.PSPath).DisplayName
    if ($displayName -like "N-able Technologies") {
        $uninstallString = (Get-ItemProperty -Path $key.PSPath).UninstallString
        $installLocation = (Get-ItemProperty -Path $key.PSPath).InstallLocation
        break
    }
}

# Check if uninstall string was found
if ($uninstallString) {
    # Remove surrounding quotes if present
    $uninstallString = $uninstallString.Trim('"')

    # Execute the uninstall command
    Start-Process -FilePath $uninstallString -ArgumentList "/qn" -Wait
    Write-Host "Uninstallation of Windows Agent completed successfully."
} else {
    Write-Host "Windows Agent not found in the registry. No uninstallation performed."
}

if ($installLocation) {
    # Remove the installation directory
    Remove-Item -Path $installLocation -Recurse -Force
    Write-Host "Installation directory removed."
} else {
    Write-Host "Installation directory not found. No removal performed."
}

# Check if reinstall switch is present
if ($Reinstall) {
    # Look in C:\Temp for the installer file that contains "WindowsAgentSetup" in the name
    $installerFile = Get-ChildItem -Path "C:\Temp\" -Filter "*WindowsAgentSetup*"
    # Reinstall the Windows Agent
    Start-Process -FilePath $installerFile.FullName -ArgumentList "/qn" -Wait
    Write-Host "Reinstallation of Windows Agent completed successfully."
}