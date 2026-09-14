# Check if the script is running with admin privileges
$isAdmin = ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)

if (-not $isAdmin) {
    # Relaunch the script with admin rights
    Write-Host "Relaunching the script as admin"
    Start-Process powershell.exe -Verb RunAs -ArgumentList "-File", $MyInvocation.MyCommand.Path
    Exit
}

# Specify the list of scripts to be executed
$scripts = @(
    "DiscoverServersV1.ps1",
    "DomainControllerDiscoveryV1.ps1",
    "ServerInfoV1.ps1",
    "ServerApplicationsV1.ps1",
    "ServerFeaturesV1.ps1",
    "ServerRolesV1.ps1",
    "CompileDocument.ps1"
)

# Specify the directory path
Write-Host "Checking to make sure C:\temp Exists, if not it will be created."
$dirPath = "C:\temp"
Start-Sleep -Milliseconds 500

# Check if the directory exists
if (-not (Test-Path -Path $dirPath)) {
    # If the directory doesn't exist, create it
    New-Item -Path $dirPath -ItemType Directory -Force
    Write-Host "Directory created at $dirPath"
}
else {
    # If the directory already exists
    Write-Host "Directory already exists at $dirPath"
}

# Set the error log file path
$errorLogPath = "C:\temp\ErrorLog.txt"

###Just for FUN

Write-Host "
______               _      ___  
| ___ \             | |    |__ \ 
| |_/ /___  __ _  __| |_   _  ) |
|    // _ \/ _` |/ _` | | | |/ / 
| |\ \  __/ (_| | (_| | |_| |_|  
\_| \_\___|\__,_|\__,_|\__, (_)  
                        __/ |    
                       |___/     "

Start-Sleep -Seconds 1

Write-Host "
 _____      _     
/  ___|    | |    
\ `--.  ___| |_   
 `--. \/ _ \ __|  
/\__/ /  __/ |_ _ 
\____/ \___|\__(_)
                  
                  "

Start-Sleep -Seconds 2

Write-Host "
 _          _           _____ _____ _ 
| |        | |         |  __ \  _  | |
| |     ___| |_ ___    | |  \/ | | | |
| |    / _ \ __/ __|   | | __| | | | |
| |___|  __/ |_\__ \   | |_\ \ \_/ /_|
\_____/\___|\__|___/    \____/\___/(_)
                                      
                                      "

# Execute each script with elevated privileges
foreach ($script in $scripts) {
    $scriptPath = Join-Path $PSScriptRoot $script

    # Execute the script and capture any errors
    $errorActionPreference = "Stop"
    try {
        Write-Host "Executing script: $script"
        & $scriptPath
        Write-Host "--------------------------------------------------------------------------"
        Write-Host "=============== script execution completed: $script ======================"
        Write-Host "--------------------------------------------------------------------------"
    }
    catch {
        # Log the error to the error log file
        $errorMessage = "Error executing script: $script`n$($_.Exception.Message)`n"
        Add-Content -Path $errorLogPath -Value $errorMessage
        Write-Host $errorMessage
    }
}

# Pause the script to keep the PowerShell window open
Write-Host "Script execution finished. Make sure to review the error log in C:\temp\ErrorLog.txt -  Press any key to exit."
$null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")