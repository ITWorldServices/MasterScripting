#Requires -Version 5.0
<#
.SYNOPSIS
    Automated installation of N-able INC Agent for Intune deployment.

.DESCRIPTION
    This script checks for an existing N-able INC Agent, and if not present or unhealthy,
    it downloads and installs the agent. The script handles token acquisition, installation,
    validation, and logging to both local and remote SQL databases.

.PARAMETER CustomerID
    The customer ID to use for the agent installation.

.PARAMETER INCServer
    The INC server to connect to (inc, inc2, inc3, inc4).

.PARAMETER LogToSQL
    Switch to enable SQL logging.
#>

[CmdletBinding()]
param(
    [Parameter()]
    [int]$CustomerID = 1618,
    
    [Parameter()]
    [ValidateSet('inc', 'inc2', 'inc3', 'inc4')]
    [string]$INCServer = 'inc2',
    
    [Parameter()]
    [switch]$LogToSQL = $false
)

#region Script Initialization
$ErrorActionPreference = "Continue"
$VerbosePreference = "Continue"

# Start transcript logging if needed
# Start-Transcript -Path "$env:ProgramData\IntuneLogs\INCAgentInstall.log" -Append

# Fix scriptName variable
$scriptName = if ($PSCommandPath) { 
    Split-Path $PSCommandPath -Leaf 
} else { 
    "Intune_AutoDeployINCAgentV2.ps1" 
}

# Create logs directory if it doesn't exist
if (!(Test-Path "$env:ProgramData\IntuneLogs")) {
    New-Item -Path "$env:ProgramData\IntuneLogs" -ItemType Directory -Force | Out-Null
}

# Constants
$INCSharedSecret = 'lkjdfjf9f890fer809fjh'
$logLevel = 'E'  # Default to Error level
Set-ExecutionPolicy -ExecutionPolicy Bypass -Scope Process -Force

# Status tracking
[system.collections.arraylist]$StatusMessages = @()

# Force TLS 1.2
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

# Normalize server name
if ($INCServer -eq 'inc1') {
    $INCServer = 'inc'
}

# Initialize Status Object
$StatusObject = [PSCustomObject]@{
    CustomerID               = $CustomerId
    RegistrationToken        = $null
    Server                   = $INCServer.ToLower()
    AgentDownloadSuccess     = $False
    TempDirectory            = [System.Environment]::GetEnvironmentVariable('TEMP', 'Machine')
    AgentDownloadPath        = "$([System.Environment]::GetEnvironmentVariable('TEMP', 'Machine'))\WindowsAgentSetup.exe"
    AgentURL                 = $null
    AgentInstallationPath    = 'C:\Program Files (x86)\N-able Technologies\Windows Agent\bin\agent.exe'
    AgentArgumentList        = $null
    CustomerIdAPICallSuccess = $False
    StatusMessage            = $null
    AgentVersion             = $null
    AgentHealth              = $null
    InstallSuccess           = $null
}
#endregion
#region Helper Functions
# Function to check agent health
function Get-AgentHealth {
    [cmdletBinding()]
    param (
        [Parameter(Mandatory)]
        [ValidateNotNullOrEmpty()]
        [string]$AgentPath
    )
    
    $result = @{
        IsHealthy = $false
        Version   = $null
        Message   = "Agent status unknown"
    }
    
    if (Test-Path $AgentPath) {
        try {
            $fileVersion = (Get-Item $AgentPath -ErrorAction Stop).VersionInfo.FileVersion
            $result.Version = $fileVersion
        } catch {
            $result.Message = "Failed to get agent version: $($_.Exception.Message)"
            return $result
        }
        
        try {
            $agentProcess = Get-Process | Where-Object { 
                $_.ProcessName -like '*agent*' -or
                $_.ProcessName -like '*Windows Agent*' -or 
                $_.ProcessName -like '*nagent*' -or
                $_.ProcessName -like '*N-able*'
            } -ErrorAction Stop
            
            if ($agentProcess) {
                $result.IsHealthy = $true
                $result.Message = "Agent is running, version: $fileVersion"
            } else {
                # Try checking service directly
                $service = Get-Service "Windows Agent Service" -ErrorAction SilentlyContinue
                if ($service -and $service.Status -eq 'Running') {
                    $result.IsHealthy = $true
                    $result.Message = "Agent service is running, version: $fileVersion"
                } else {
                    $result.Message = "Agent file exists but process/service is not running"
                }
            }
        } catch {
            $result.Message = "Error checking agent health: $($_.Exception.Message)"
        }
    } else {
        $result.Message = "Agent installation path not found: $AgentPath"
    }
    
    return $result
}


function Write-ImpactMessage {
    [cmdletBinding()]
    param(
        [ValidateSet('I', 'W', 'E')]
        [string]$loglevel,
        [string]$scriptname,
        [string]$message,
        [int]$CustomerId,
        [int]$DeviceID,
        [string]$NCentralAssetTag,
        [ValidateSet('inc', 'inc2', 'inc3', 'inc4')] # I forgot what I put this in here for.  Not useable at the moment.  1/22/25 JB
        [string]$INCServer,
        [switch]$UseDevSQLTable
    )

    BEGIN {            
        
        
        Function Get-INCAgentCustomerID {
            [cmdletbinding()]
            Param()
            BEGIN {
                $XMLFilePath = 'C:\Program Files (x86)\N-Able Technologies\Windows Agent\config\AgentMaintenanceSchedules.xml'
            }
            PROCESS {
                # Check if the AgentMaintenanceSchedules.xml file exists.
                if (Test-Path -Path $XMLFilePath) {
                    try {
                        # Import the XML file into a PowerShell object.
                        $xml = [xml](Get-Content -Path $XMLFilePath -ErrorAction STOP)
                    }
                    catch {
                        Write-Error "Failed to read or parse the XML file '$($XMLFilePath)'"
                    }
                }
                else {
                    Write-Error "The file '$($XMLFilePath)' does not exist."
                }

                # Check if the $xml variable is not null or empty and if the RebootMessageLogoURL property exists.
                if ($xml -and $xml.AgentMaintenanceSchedules.RebootMessageLogoURL) {
                    try {
                        # Extract the line containing the Customer ID.
                        $line = $xml.AgentMaintenanceSchedules.RebootMessageLogoURL
           
                        # Extract the Customer ID from the line and store it into a variable.
                        $CustomerID = [regex]::Match($line, '\d+').Value
                        $Ok = $True
                    }
                    catch {
                        if (-not $CustomerID) {
                            Write-Error "Error parsing '$($XMLFilePath)' file for the CustomerID"
                        }
                    }
                    if ($OK) {
                        $CustomerID
                    }
                }
                else {
                    Write-Error 'The XML content at '$($XMLFilePath)' is invalid or the RebootMessageLogoURL property does not exist.'
                }
            }
            END {}
        }
        
        Function Get-INCAgentAssetTag {
            [cmdletbinding()]
            Param()
            BEGIN {
                $XMLFilePath = 'C:\Program Files (x86)\N-Able Technologies\NcentralAsset.xml'
            }
            PROCESS {
                # Check if the AgentMaintenanceSchedules.xml file exists.
                if (Test-Path -Path $XMLFilePath) {
                    try {
                        # Import the XML file into a PowerShell object.
                        $xml = [xml](Get-Content -Path $XMLFilePath -ErrorAction STOP)
                    }
                    catch {
                        Write-Error "Failed to read or parse the XML file '$($XMLFilePath)'"
                    }
                }
                else {
                    Write-Error "The file '$($XMLFilePath)' does not exist."
                }

                if ($xml -and $xml.ncentralasset.ncentralassettag) {
                    $ncentralassettag = $xml.ncentralasset.ncentralassettag
                }
                else {
                    Write-Error "The XML content is invalid or the Applianceconfig property does not exist in file '$($XMLFilePath)'"
                }
                if ($ncentralassettag) {
                    $ncentralassettag
                }
            }
            END {}
        }

        Function Get-INCAgentDeviceID {
            [cmdletbinding()]
            Param()
            BEGIN {
                $XMLFilePath = 'C:\Program Files (x86)\N-Able Technologies\Windows Agent\config\executionerConfig.xml'
            }
            PROCESS {
                # Check if the executionerConfig.xml file exists.
                if (Test-Path -Path $XMLFilePath) {
                    try {
                        # Import the XML file into a PowerShell object.
                        $xml = [xml] (Get-Content -Path $XMLFilePath -ErrorAction STOP)
                    }
                    catch {
                        Write-Error "Failed to read or parse the XML file '$($XMLFilePath)'"
                    }
                }
                else {
                    Write-Error "The file '$($XMLFilePath)' does not exist"
                }

                # Check if the $xml variable is not null or empty.
                if ($xml) {
                    try {
                        # Convert the XML object to a string and search for the agentID pattern.
                        $xmlString = $xml.OuterXml
                        $match = [regex]::Match($xmlString, 'agentid=\d+')
            
                        if ($match.Success) {
                            # Extract the "AgentID='number value' from the match and store it into a variable.
                            $line = $match.Value
                            # Extract the number value from the line and store it into a variable.
                            $DeviceID = -join ($line -split '\D+')
                            $OK = $true
                        }
                        else {
                            Write-Error "Found XML file '$($XMLFilePath)' but no DeviceID found in the XML content"
                        }
                    }
                    catch {
                        Write-Error "An error occurred while searching for the DeviceID: $($Error[0])"
                    }
                    if ($OK) {
                        $DeviceID
                    }
                }
                else {
                    Write-Error "The XML content was not found or is empty in the file '$($XMLFilePath)'"
                }
            }
            END {}
        }
    }

    PROCESS {
        if (!$PSBoundParameters.ContainsKey('CustomerId')) {
            Try {
                $CustomerId = Get-INCAgentCustomerID -ErrorAction STOP
            }
            Catch {}
        }
        if (!$PSBoundParameters.ContainsKey('DeviceID')) {
            Try {
                $DeviceID = Get-INCAgentDeviceID -ErrorAction STOP
            }
            Catch {}
        }
        if (!$PSBoundParameters.ContainsKey('NCentralAssetTag')) {
            Try {
                $NCentralAssetTag = Get-INCAgentAssetTag -ErrorAction STOP
            }
            Catch {}
        }
       
        Try {
            $MessageHash = @{
                ScriptName       = $ScriptName
                LogLevel         = $LogLevel
                DeviceName       = $env:COMPUTERNAME
                DeviceID         = $DeviceID
                NcentralAssetTag = $NcentralAssetTag
                CustomerID       = $CustomerID
                Message          = $Message
                sharedSecret     = '089324hjweaasdfasdfsdfa433sdfas'
            }

            $json_body = ConvertTo-Json $MessageHash
            Write-Verbose $($Json_Body)
            $RequestResults = Invoke-WebRequest 'https://sortarius.impactnfr.com/MIT_AgentLog' -Body $json_body -Method POST -UseBasicParsing -ErrorAction STOP
        }
        Catch {
            Write-Verbose -Message "Write-ImpactMessage failed, details: $($_.Exception.message)"
            Write-Verbose $($RequestResults)
        }
    }
    END {}
}

Function WV {
    Param(
        [Parameter()]
        [string]$prefix,
        [Parameter(Mandatory = $true)]
        [string]$message,
        [Parameter()]
        [string]$File
    )
    $time = Get-Date -f MM-dd-HH:mm:ss:ffff
    $prefixString = if ($prefix) { "[$($prefix.PadRight(10,' '))]" } else { "[SCRIPT]" }
    
    # Always output to verbose stream
    Write-Verbose "$time $prefixString $message"
    
    # Also log to file if specified
    if ($File) {
        try {
            [pscustomobject]@{
                Time    = $time
                Prefix  = $Prefix
                Message = $Message
            } | Export-Csv -Path $File -Append -NoTypeInformation -ErrorAction Stop
        }
        catch {
            Write-Verbose "Failed to write to log file ${File}: $($_.Exception.Message)"
        }
    }
}

function New-WebClientRequest {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$URL,
        [Parameter(Mandatory = $true)]
        [string]$SavePath,
        [switch]$CalculateFileSize
    )
    
    $data = @{
        success       = $false
        statusMessage = [System.Collections.ArrayList]@()
    }

    try {
        WV -prefix $ScriptName -Message "Attempting to download from: $URL"
        WV -prefix $ScriptName -Message "Saving to: $SavePath"
        
        if ($PSBoundParameters.ContainsKey('CalculateFileSize')) {
            # Test web request first
            $testRequest = [System.Net.WebRequest]::Create($URL)
            $testRequest.Method = 'HEAD'
            try {
                $testResponse = $testRequest.GetResponse()
                $fileSize = $testResponse.ContentLength
                $null = $data.statusMessage.Add("File size on server: $($fileSize/1MB) MB")
                $testResponse.Close()
            }
            catch {
                $null = $data.statusMessage.Add("Failed to get file info: $($_.Exception.Message)")
                throw
            }
        }
        # Proceed with download
        $WebClient = New-Object System.Net.WebClient
        $WebClient.DownloadFile($URL, $SavePath)
        
        if (Test-Path $SavePath) {
            $downloadedSize = (Get-Item $SavePath).Length
            $null = $data.statusMessage.Add("Downloaded file size: $($downloadedSize/1MB) MB")
            if ($downloadedSize -gt 0) {
                $data.success = $true
                $null = $data.statusMessage.Add("Download success: $($SavePath)")
            }
            else {
                $null = $data.statusMessage.Add('Downloaded file is empty')
                $data.success = $false
            }
        }
        else {
            $null = $data.statusMessage.Add('File not found after download')
            $data.success = $false
        }
    }
    catch {
        $errorMessage = if ($_.Exception.InnerException) {
            $_.Exception.InnerException.Message
        }
        else {
            $_.Exception.Message
        }
        $null = $data.statusMessage.Add("Download Error: $($errorMessage)")
        $data.success = $false
    }
    $data
}

#endregion

#region Main Script Logic
WV -prefix $ScriptName -Message 'Starting agent health check...'
Try {
    $agentHealth = Get-AgentHealth -AgentPath $StatusObject.AgentInstallationPath -ErrorAction STOP
    $StatusObject.AgentHealth = $agentHealth.IsHealthy
    $StatusObject.AgentVersion = $agentHealth.Version
}
Catch {
    WV -prefix $ScriptName -Message "Error checking agent health: $($_.Exception.Message)"
}

if (-not $agentHealth.IsHealthy) {
    $LogToSQL = $true
    $null = $StatusMessages.Add("INC Agent needs installation: $($agentHealth.Message)")
    WV -prefix $ScriptName -Message 'INC Agent needs installation. Starting installation process...'
    
    # Get registration token from API
    WV -prefix $ScriptName -Message 'Requesting registration token...'
    $RequestObject = @{
        CustomerID   = $StatusObject.CustomerId
        Server       = $StatusObject.Server
        SharedSecret = $INCSharedSecret
    }

    try {
        $CustomerObject = Invoke-RestMethod 'https://sortarius.impactnfr.com/INCCustomerId' -Method POST -Body ($RequestObject | ConvertTo-Json) -ErrorAction STOP
        $StatusObject.CustomerIdAPICallSuccess = $true
        $StatusObject.RegistrationToken = $CustomerObject.RegistrationToken
        $StatusObject.Server = $CustomerObject.Server.ToLower()
        $StatusObject.AgentURL = "https://$($StatusObject.Server).impactnetworking.com/download/current/winnt/N-central/WindowsAgentSetup.exe"
        
        WV -prefix $ScriptName -Message 'Successfully obtained registration token'
        $null = $StatusMessages.Add("Got Registration Token $($StatusObject.RegistrationToken) for CustomerId $($StatusObject.CustomerID)")
        $null = $StatusMessages.Add("Download URL: $($StatusObject.AgentURL)")
    }
    catch {
        $StatusObject.CustomerIdAPICallSuccess = $false
        $errorMessage = "Failed to get Registration Token: $($_.Exception.Message)"
        $null = $StatusMessages.Add($errorMessage)
        WV -prefix $ScriptName -Message "$($errorMessage)"
    }

    # Download and install agent
    if ($StatusObject.RegistrationToken -and $StatusObject.Server -and $StatusObject.CustomerId) {
        WV -prefix $ScriptName -Message 'Valid Registration Token and Server found. Proceeding with agent installation...'
    
        # Region Download
        Try {
            $downloadResult = New-WebClientRequest -URL $StatusObject.AgentURL -SavePath $StatusObject.AgentDownloadPath -ErrorAction STOP
        }
        Catch {
            WV -prefix $ScriptName -Message "Agent failed to download: $($Error[0])"
        }
        $StatusObject.AgentDownloadSuccess = $downloadResult.success        

        # Add download status messages
        foreach ($msg in $downloadResult.statusMessage) {
            $null = $StatusMessages.Add($msg)
        }

        if ($StatusObject.AgentDownloadSuccess -and (Test-Path $StatusObject.AgentDownloadPath)) {
            # Validate the installer before running
            $fileInfo = Get-Item $StatusObject.AgentDownloadPath
            if ($fileInfo.Length -gt 1MB) {  # Sanity check - installer should be bigger than 1MB
                # Prepare the arguments properly
                $arguments = @(
                    '/s',
                    '/v"',
                    '/qn',
                    "CUSTOMERID=$($StatusObject.CustomerId)",
                    'CUSTOMERSPECIFIC=1',
                    "REGISTRATION_TOKEN=$($StatusObject.RegistrationToken)",
                    'SERVERPROTOCOL=HTTPS',
                    "SERVERADDRESS=$($StatusObject.Server).impactnetworking.com",
                    'SERVERPORT=443"'
                )
                
                $argumentString = $arguments -join ' '
                WV -prefix $ScriptName -Message "Starting installation with arguments: $argumentString"
                
                try {
                    # Run the installer with a timeout
                    $process = Start-Process -FilePath $StatusObject.AgentDownloadPath -ArgumentList $arguments -PassThru -WindowStyle Hidden -ErrorAction Stop
                    
                    # Wait for completion with timeout
                    $timeoutSeconds = 300  # 5 minutes
                    $process.WaitForExit($timeoutSeconds * 1000)
                    
                    if (!$process.HasExited) {
                        WV -prefix $ScriptName -Message "Installation timeout after $timeoutSeconds seconds"
                        try {
                            $process.Kill()
                        } catch {
                            WV -prefix $ScriptName -Message "Failed to kill installer process"
                        }
                        $StatusObject.InstallSuccess = $false
                    } else {
                        WV -prefix $ScriptName -Message "Installation completed with exit code: $($process.ExitCode)"
                        
                        # Wait for service to start
                        $serviceTimeout = 60  # 1 minute - increased for more startup time
                        $startTime = Get-Date
                        $serviceStarted = $false
                        
                        while ((Get-Date) -lt ($startTime.AddSeconds($serviceTimeout))) {
                            if (Get-Service 'Windows Agent Service' -ErrorAction SilentlyContinue) {
                                $serviceStarted = $true
                                break
                            }
                            Start-Sleep -Seconds 5
                        }
                        
                        if ($serviceStarted) {
                            WV -prefix $ScriptName -Message "Agent service found"
                            
                            # Try to start service explicitly if it exists but isn't running
                            $service = Get-Service 'Windows Agent Service' -ErrorAction SilentlyContinue
                            if ($service -and $service.Status -ne 'Running') {
                                WV -prefix $ScriptName -Message "Starting agent service..."
                                try {
                                    Start-Service 'Windows Agent Service' -ErrorAction SilentlyContinue
                                    Start-Sleep -Seconds 10
                                } catch {
                                    WV -prefix $ScriptName -Message "Error starting service: $($_.Exception.Message)"
                                }
                            }
                            
                            $StatusObject.InstallSuccess = $true
                            
                            # Verify installation - wait longer for agent to fully initialize
                            WV -prefix $ScriptName -Message 'Verifying installation...'
                            Start-Sleep -Seconds 30
                            Try {
                                $postInstallHealth = Get-AgentHealth -AgentPath $StatusObject.AgentInstallationPath -ErrorAction STOP
                            }
                            Catch {
                                $null = $StatusMessages.Add("Error checking agent health after installation: $($Error[0])")
                                WV -prefix $ScriptName -Message "Error checking agent health after installation: $($Error[0])"
                            }
                            
                            if ($postInstallHealth.IsHealthy) {
                                $StatusObject.AgentHealth = $true
                                $StatusObject.InstallSuccess = $true
                                $StatusObject.AgentVersion = $postInstallHealth.Version
                                $null = $StatusMessages.Add('Post-install verification: Agent is healthy and running')
                                WV -prefix $ScriptName -Message 'Post-install verification: Agent is healthy and running'
                            } else {
                                $null = $StatusMessages.Add("Post-install verification failed: $($postInstallHealth.Message)")
                                WV -prefix $ScriptName -Message "Post-install verification failed: $($postInstallHealth.Message)"
                            }
                            
                            # Cleanup installer file
                            Try {
                                Remove-Item $StatusObject.AgentDownloadPath -Force -ErrorAction STOP
                                $null = $StatusMessages.Add('Cleaned up installer file after installation')
                            }
                            Catch {
                                $null = $StatusMessages.Add('Failed to clean up installer file after installation')
                            }
                        } else {
                            WV -prefix $ScriptName -Message "Timed out waiting for agent service"
                            $StatusObject.InstallSuccess = $false
                        }
                    }
                } catch {
                    WV -prefix $ScriptName -Message "Failed to start installer: $($_.Exception.Message)"
                    $StatusObject.InstallSuccess = $false
                }
            } else {
                WV -prefix $ScriptName -Message "Installer file is too small or corrupt: $($fileInfo.Length) bytes"
                $StatusObject.InstallSuccess = $false
            }
        } else {
            WV -prefix $ScriptName -Message "Cannot install - download failed or file missing"
            $StatusObject.InstallSuccess = $false
        }
    } else {
        WV -prefix $ScriptName -Message "Missing required installation parameters"
        $null = $StatusMessages.Add("Missing required installation parameters")
    }
} else {
    WV -prefix $ScriptName -Message 'INC Agent is healthy'
    $null = $StatusMessages.Add("INC Agent is healthy: $($agentHealth.Message)")
}
#endregion

#region Finalization and Logging
# Final status update
$StatusObject.StatusMessage = $StatusMessages

if ($StatusObject.AgentHealth -eq $true) {
    $LogLevel = 'I'
} else {
    $LogLevel = 'E'
}


# Final status update
$StatusObject.StatusMessage = $StatusMessages
$StatusObjectJSON = $StatusObject | ConvertTo-Json -Depth 10

if($StatusObject.AgentHealth -eq $true) {
    $LogToSQL = $true
    $logLevel = 'I'
    WV -prefix $ScriptName -Message "Installation successful"
    if ($LogToSQL) {
        Write-ImpactMessage `
            -LogLevel $logLevel `
            -ScriptName "Install Windows INC Agent v2" `
            -Message $StatusObjectJSON `
            -UseDevSQLTable
        WV -prefix $ScriptName -Message "Logged to SQL"
    }
    #Stop-Transcript
    exit 0 
} else {
    $LogToSQL = $true
    $logLevel = 'E'
    WV -prefix $ScriptName -Message "Installation failed or agent not healthy"
    if ($LogToSQL) {
        Write-ImpactMessage `
            -LogLevel $logLevel `
            -ScriptName "Install Windows INC Agent v2" `
            -Message $StatusObjectJSON `
            -UseDevSQLTable
        WV -prefix $ScriptName -Message "Logged to SQL"
    }
    #Stop-Transcript
    exit 1
}

# Return status object for further processing if needed
$StatusObject
#endregion