# Enhanced Service Account Report Script
# Comprehensive security analysis of service accounts across multiple servers
# Basic usage
#.\ServiceAccountsReport.ps1

# Advanced usage with all features
#.\ServiceAccountsReport.ps1 -ServerListPath "prod_servers.txt" -OutputPath ".\SecurityReports" -EnableADLookup -IncludeNonAdminServices -MaxParallelJobs 10

# Quick scan without AD lookup
#.\ServiceAccountsReport.ps1 -MaxParallelJobs 3
param(
    [Parameter(Mandatory=$false)]
    [string]$ServerListPath = "servers.txt",
    
    [Parameter(Mandatory=$false)]
    [string]$OutputPath = ".\Reports",
    
    [Parameter(Mandatory=$false)]
    [switch]$IncludeNonAdminServices,
    
    [Parameter(Mandatory=$false)]
    [int]$TimeoutSeconds = 30,
    
    [Parameter(Mandatory=$false)]
    [int]$MaxParallelJobs = 5,
    
    [Parameter(Mandatory=$false)]
    [switch]$EnableADLookup
)

# Validate input file exists
if (-not (Test-Path $ServerListPath)) {
    Write-Error "Server list file not found at: $ServerListPath"
    exit 1
}

# Create output directory if it doesn't exist
if (-not (Test-Path $OutputPath)) {
    New-Item -ItemType Directory -Path $OutputPath | Out-Null
}

# Check for Active Directory module if AD lookup is enabled
if ($EnableADLookup) {
    if (-not (Get-Module -ListAvailable -Name ActiveDirectory)) {
        Write-Warning "Active Directory module not available. Disabling AD lookups."
        $EnableADLookup = $false
    } else {
        Import-Module ActiveDirectory -ErrorAction SilentlyContinue
    }
}

# Define list of servers
$servers = Get-Content -Path $ServerListPath | Where-Object { $_ -match '\S' } | Sort-Object -Unique

# Define patterns for administrator accounts
$adminPatterns = @(
    ".*admin.*",
    ".*administrator.*",
    ".*svc.*",
    ".*service.*",
    ".*local.*system",
    ".*nt authority.*system",
    ".*network service",
    ".*local service"
)

# Define suspicious patterns for security analysis
$suspiciousPatterns = @{
    HighPrivilegeAccounts = @("domain.*admin", "enterprise.*admin", "schema.*admin")
    SuspiciousPaths = @("temp", "users", "downloads", "recycle", "appdata")
    UnusualExtensions = @("\.tmp", "\.bat", "\.vbs", "\.ps1", "\.cmd")
}

# Function to test server connectivity
function Test-ServerConnectivity {
    param($Server, $TimeoutSeconds = 30)
    try {
        $ping = Test-Connection -ComputerName $Server -Count 1 -Quiet
        if ($ping) {
            $wmi = Get-WmiObject -Class Win32_OperatingSystem -ComputerName $Server -ErrorAction Stop
            return $true
        }
        return $false
    }
    catch {
        return $false
    }
}

# Function to get AD account details
function Get-ADAccountDetails {
    param($AccountName)
    
    if (-not $EnableADLookup) {
        return [PSCustomObject]@{
            IsServiceAccount = "N/A"
            IsDisabled = "N/A"
            LastLogon = "N/A"
            PasswordAge = "N/A"
            MemberOf = "N/A"
        }
    }
    
    try {
        # Clean up the account name (remove domain prefix if present)
        $cleanAccountName = $AccountName -replace ".*\\", ""
        
        # Skip system accounts
        if ($cleanAccountName -match "^(LocalSystem|NetworkService|LocalService|NT AUTHORITY)") {
            return [PSCustomObject]@{
                IsServiceAccount = $true
                IsDisabled = $false
                LastLogon = "System Account"
                PasswordAge = "System Account"
                MemberOf = "System Account"
            }
        }
        
        $adUser = Get-ADUser -Identity $cleanAccountName -Properties LastLogonDate, PasswordLastSet, MemberOf -ErrorAction Stop
        
        return [PSCustomObject]@{
            IsServiceAccount = $false
            IsDisabled = -not $adUser.Enabled
            LastLogon = $adUser.LastLogonDate
            PasswordAge = if ($adUser.PasswordLastSet) { [math]::Round(((Get-Date) - $adUser.PasswordLastSet).TotalDays, 0) } else { "Unknown" }
            MemberOf = ($adUser.MemberOf | Get-ADGroup | Select-Object -ExpandProperty Name) -join "; "
        }
    }
    catch {
        return [PSCustomObject]@{
            IsServiceAccount = $false
            IsDisabled = "Unknown"
            LastLogon = "Unknown"
            PasswordAge = "Unknown"
            MemberOf = "Account Not Found"
        }
    }
}

# Function for security analysis
function Get-SecurityAnalysis {
    param($ServiceDetails)
    
    $securityFlags = @()
    
    # Check for high-privilege accounts
    foreach ($pattern in $suspiciousPatterns.HighPrivilegeAccounts) {
        if ($ServiceDetails.StartName -match $pattern) {
            $securityFlags += "HIGH_PRIVILEGE_ACCOUNT"
        }
    }
    
    # Check for suspicious paths
    foreach ($pattern in $suspiciousPatterns.SuspiciousPaths) {
        if ($ServiceDetails.PathName -match $pattern) {
            $securityFlags += "SUSPICIOUS_PATH"
        }
    }
    
    # Check for services running from user directories
    if ($ServiceDetails.PathName -match "C:\\Users\\") {
        $securityFlags += "USER_DIRECTORY_EXECUTION"
    }
    
    # Check for unusual file extensions
    foreach ($pattern in $suspiciousPatterns.UnusualExtensions) {
        if ($ServiceDetails.PathName -match $pattern) {
            $securityFlags += "UNUSUAL_EXECUTABLE"
        }
    }
    
    # Check if service is running but disabled
    if ($ServiceDetails.State -eq "Running" -and $ServiceDetails.StartMode -eq "Disabled") {
        $securityFlags += "RUNNING_WHILE_DISABLED"
    }
    
    return if ($securityFlags.Count -gt 0) { $securityFlags -join "; " } else { "None" }
}

# Function for CIS compliance checks
function Test-CISCompliance {
    param($ServiceDetails)
    
    $cisFindings = @()
    
    # CIS Control 5.1: Ensure unnecessary services are disabled
    $unnecessaryServices = @("Telnet", "SNMP", "Spooler", "Fax", "RemoteRegistry")
    if ($unnecessaryServices -contains $ServiceDetails.ServiceName -and $ServiceDetails.State -eq "Running") {
        $cisFindings += "CIS-5.1: Unnecessary service running"
    }
    
    # CIS Control 9.1: Ensure services run with least privilege
    if ($ServiceDetails.StartName -match "administrator|admin" -and $ServiceDetails.StartName -notmatch "local.*service|network.*service|nt authority") {
        $cisFindings += "CIS-9.1: Service running with excessive privileges"
    }
    
    # Check for services with manual start mode that are running
    if ($ServiceDetails.StartMode -eq "Manual" -and $ServiceDetails.State -eq "Running") {
        $cisFindings += "CIS-INFO: Manual service currently running"
    }
    
    return if ($cisFindings.Count -gt 0) { $cisFindings -join "; " } else { "Compliant" }
}

# Function to get detailed service information
function Get-ServiceDetails {
    param($Server, $Service)
    
    try {
        # Get process details if available
        $processInfo = $null
        if ($Service.ProcessId -and $Service.ProcessId -gt 0) {
            $processInfo = Get-WmiObject -Class Win32_Process -ComputerName $Server -Filter "ProcessId=$($Service.ProcessId)" -ErrorAction SilentlyContinue
        }
        
        # Create base service object
        $serviceObj = [PSCustomObject]@{
            ServerName     = $Server
            ServiceName    = $Service.Name
            DisplayName    = $Service.DisplayName
            StartMode      = $Service.StartMode
            State          = $Service.State
            StartName      = $Service.StartName
            Description    = $Service.Description
            PathName       = $Service.PathName
            ProcessId      = $Service.ProcessId
            BinaryPath     = if ($processInfo) { $processInfo.ExecutablePath } else { "N/A" }
            CompanyName    = if ($processInfo) { $processInfo.Company } else { "N/A" }
            MemoryUsageMB  = if ($processInfo -and $processInfo.WorkingSetSize) { [math]::Round($processInfo.WorkingSetSize / 1MB, 2) } else { 0 }
        }
        
        # Add security analysis
        $serviceObj | Add-Member -NotePropertyName "SecurityFlags" -NotePropertyValue (Get-SecurityAnalysis $serviceObj)
        
        # Add CIS compliance check
        $serviceObj | Add-Member -NotePropertyName "CISFindings" -NotePropertyValue (Test-CISCompliance $serviceObj)
        
        # Add AD account details
        $adDetails = Get-ADAccountDetails $serviceObj.StartName
        $serviceObj | Add-Member -NotePropertyName "IsDisabled" -NotePropertyValue $adDetails.IsDisabled
        $serviceObj | Add-Member -NotePropertyName "LastLogon" -NotePropertyValue $adDetails.LastLogon
        $serviceObj | Add-Member -NotePropertyName "PasswordAgeDays" -NotePropertyValue $adDetails.PasswordAge
        $serviceObj | Add-Member -NotePropertyName "MemberOf" -NotePropertyValue $adDetails.MemberOf
        
        return $serviceObj
    }
    catch {
        Write-Warning "Failed to get detailed information for service $($Service.Name) on $Server: $_"
        return $null
    }
}

# Function for parallel server scanning
function Start-ParallelServerScan {
    param($Servers, $MaxJobs = 5)
    
    $jobs = @()
    $allResults = @()
    
    Write-Host "Starting parallel scan with up to $MaxJobs concurrent jobs..." -ForegroundColor Yellow
    
    foreach ($server in $Servers) {
        # Wait if we have too many jobs running
        while ((Get-Job -State Running).Count -ge $MaxJobs) {
            Start-Sleep -Seconds 1
        }
        
        # Start a new job for this server
        $job = Start-Job -ScriptBlock {
            param($ServerName, $AdminPatterns, $IncludeNonAdminServices, $EnableADLookup, $SuspiciousPatterns)
            
            # Define functions within the job scope
            function Get-SecurityAnalysis {
                param($ServiceDetails)
                $securityFlags = @()
                foreach ($pattern in $SuspiciousPatterns.HighPrivilegeAccounts) {
                    if ($ServiceDetails.StartName -match $pattern) { $securityFlags += "HIGH_PRIVILEGE_ACCOUNT" }
                }
                foreach ($pattern in $SuspiciousPatterns.SuspiciousPaths) {
                    if ($ServiceDetails.PathName -match $pattern) { $securityFlags += "SUSPICIOUS_PATH" }
                }
                if ($ServiceDetails.PathName -match "C:\\Users\\") { $securityFlags += "USER_DIRECTORY_EXECUTION" }
                foreach ($pattern in $SuspiciousPatterns.UnusualExtensions) {
                    if ($ServiceDetails.PathName -match $pattern) { $securityFlags += "UNUSUAL_EXECUTABLE" }
                }
                if ($ServiceDetails.State -eq "Running" -and $ServiceDetails.StartMode -eq "Disabled") {
                    $securityFlags += "RUNNING_WHILE_DISABLED"
                }
                return if ($securityFlags.Count -gt 0) { $securityFlags -join "; " } else { "None" }
            }
            
            function Test-CISCompliance {
                param($ServiceDetails)
                $cisFindings = @()
                $unnecessaryServices = @("Telnet", "SNMP", "Spooler", "Fax", "RemoteRegistry")
                if ($unnecessaryServices -contains $ServiceDetails.ServiceName -and $ServiceDetails.State -eq "Running") {
                    $cisFindings += "CIS-5.1: Unnecessary service running"
                }
                if ($ServiceDetails.StartName -match "administrator|admin" -and $ServiceDetails.StartName -notmatch "local.*service|network.*service|nt authority") {
                    $cisFindings += "CIS-9.1: Service running with excessive privileges"
                }
                if ($ServiceDetails.StartMode -eq "Manual" -and $ServiceDetails.State -eq "Running") {
                    $cisFindings += "CIS-INFO: Manual service currently running"
                }
                return if ($cisFindings.Count -gt 0) { $cisFindings -join "; " } else { "Compliant" }
            }
            
            try {
                $services = Get-WmiObject -Class Win32_Service -ComputerName $ServerName -ErrorAction Stop
                $serverResults = @()
                
                foreach ($service in $services) {
                    $isAdminAccount = $false
                    foreach ($pattern in $AdminPatterns) {
                        if ($service.StartName -match $pattern) {
                            $isAdminAccount = $true
                            break
                        }
                    }
                    
                    if ($isAdminAccount -or $IncludeNonAdminServices) {
                        # Get process details if available
                        $processInfo = $null
                        if ($service.ProcessId -and $service.ProcessId -gt 0) {
                            $processInfo = Get-WmiObject -Class Win32_Process -ComputerName $ServerName -Filter "ProcessId=$($service.ProcessId)" -ErrorAction SilentlyContinue
                        }
                        
                        $serviceObj = [PSCustomObject]@{
                            ServerName     = $ServerName
                            ServiceName    = $service.Name
                            DisplayName    = $service.DisplayName
                            StartMode      = $service.StartMode
                            State          = $service.State
                            StartName      = $service.StartName
                            Description    = $service.Description
                            PathName       = $service.PathName
                            ProcessId      = $service.ProcessId
                            BinaryPath     = if ($processInfo) { $processInfo.ExecutablePath } else { "N/A" }
                            CompanyName    = if ($processInfo) { $processInfo.Company } else { "N/A" }
                            MemoryUsageMB  = if ($processInfo -and $processInfo.WorkingSetSize) { [math]::Round($processInfo.WorkingSetSize / 1MB, 2) } else { 0 }
                            SecurityFlags  = (Get-SecurityAnalysis $serviceObj)
                            CISFindings    = (Test-CISCompliance $serviceObj)
                            IsDisabled     = "N/A"
                            LastLogon      = "N/A"
                            PasswordAgeDays = "N/A"
                            MemberOf       = "N/A"
                        }
                        
                        $serverResults += $serviceObj
                    }
                }
                
                return $serverResults
            }
            catch {
                return @{ Error = $_.Exception.Message; Server = $ServerName }
            }
        } -ArgumentList $server, $adminPatterns, $IncludeNonAdminServices, $EnableADLookup, $suspiciousPatterns
        
        $jobs += $job
    }
    
    # Wait for all jobs to complete and collect results
    $completedJobs = 0
    foreach ($job in $jobs) {
        $completedJobs++
        Write-Progress -Activity "Collecting Results" -Status "Processing job $completedJobs of $($jobs.Count)" -PercentComplete (($completedJobs / $jobs.Count) * 100)
        
        $jobResult = Receive-Job -Job $job -Wait
        Remove-Job -Job $job
        
        if ($jobResult -and $jobResult.Error) {
            Write-Warning "Failed to scan $($jobResult.Server): $($jobResult.Error)"
            Add-Content -Path "$OutputPath\ServiceScanErrors.log" -Value "$(Get-Date): Failed to scan $($jobResult.Server) - $($jobResult.Error)"
        } elseif ($jobResult) {
            $allResults += $jobResult
        }
    }
    
    Write-Progress -Activity "Collecting Results" -Completed
    return $allResults
}

# Function for sequential server scanning
function Start-SequentialServerScan {
    param($Servers)
    
    $allResults = @()
    $totalServers = $Servers.Count
    $currentServer = 0
    
    foreach ($server in $Servers) {
        $currentServer++
        Write-Progress -Activity "Scanning Servers" -Status "Processing $server" -PercentComplete (($currentServer / $totalServers) * 100)
        
        Write-Host "Scanning $server... ($currentServer of $totalServers)" -ForegroundColor Cyan
        
        if (-not (Test-ServerConnectivity $server $TimeoutSeconds)) {
            Write-Warning "Server $server is not reachable"
            Add-Content -Path "$OutputPath\ServiceScanErrors.log" -Value "$(Get-Date): Server $server is not reachable"
            continue
        }
        
        try {
            $services = Get-WmiObject -Class Win32_Service -ComputerName $server -ErrorAction Stop
            
            foreach ($service in $services) {
                $isAdminAccount = $false
                foreach ($pattern in $adminPatterns) {
                    if ($service.StartName -match $pattern) {
                        $isAdminAccount = $true
                        break
                    }
                }
                
                if ($isAdminAccount -or $IncludeNonAdminServices) {
                    $serviceDetails = Get-ServiceDetails -Server $server -Service $service
                    if ($serviceDetails) {
                        $allResults += $serviceDetails
                    }
                }
            }
        }
        catch {
            Write-Warning "Failed to connect to $server: $_"
            Add-Content -Path "$OutputPath\ServiceScanErrors.log" -Value "$(Get-Date): Failed to scan $server - $_"
        }
    }
    
    Write-Progress -Activity "Scanning Servers" -Completed
    return $allResults
}

# Function to export enhanced reports
function Export-EnhancedReports {
    param($Results, $OutputPath)
    
    $timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
    
    # Sort results
    $sortedResults = $Results | Sort-Object ServerName, ServiceName
    
    # Export CSV
    $csvPath = "$OutputPath\ServiceAccountsReport_$timestamp.csv"
    $sortedResults | Export-Csv -Path $csvPath -NoTypeInformation -Encoding UTF8
    
    # Generate HTML report
    $highRiskServices = $sortedResults | Where-Object { $_.SecurityFlags -ne "None" }
    $nonCompliantServices = $sortedResults | Where-Object { $_.CISFindings -ne "Compliant" }
    $disabledAccountServices = $sortedResults | Where-Object { $_.IsDisabled -eq $true }
    
    $htmlReport = @"
<!DOCTYPE html>
<html>
<head>
    <title>Enhanced Service Account Security Report</title>
    <style>
        body { font-family: Arial, sans-serif; margin: 20px; }
        table { border-collapse: collapse; width: 100%; margin: 20px 0; }
        th, td { border: 1px solid #ddd; padding: 8px; text-align: left; }
        th { background-color: #f2f2f2; font-weight: bold; }
        tr:nth-child(even) { background-color: #f9f9f9; }
        .high-risk { background-color: #ffebee; }
        .warning { color: #ff9800; }
        .error { color: #f44336; }
        .success { color: #4caf50; }
        .summary-box { background-color: #e3f2fd; padding: 15px; border-radius: 5px; margin: 20px 0; }
        .risk-high { background-color: #ffcdd2; }
        .risk-medium { background-color: #fff3e0; }
        .risk-low { background-color: #e8f5e8; }
    </style>
</head>
<body>
    <h1>Enhanced Service Account Security Report</h1>
    <div class="summary-box">
        <h2>Executive Summary</h2>
        <p><strong>Generated:</strong> $(Get-Date)</p>
        <ul>
            <li><strong>Total Servers Scanned:</strong> $(($sortedResults | Select-Object -Unique ServerName).Count)</li>
            <li><strong>Total Services Found:</strong> $($sortedResults.Count)</li>
            <li><strong>High Risk Services:</strong> <span class="error">$($highRiskServices.Count)</span></li>
            <li><strong>Non-Compliant Services:</strong> <span class="warning">$($nonCompliantServices.Count)</span></li>
            <li><strong>Services with Disabled Accounts:</strong> <span class="warning">$($disabledAccountServices.Count)</span></li>
        </ul>
    </div>

    <h2>High Risk Services</h2>
    <table>
        <tr>
            <th>Server</th>
            <th>Service Name</th>
            <th>Account</th>
            <th>Security Flags</th>
            <th>State</th>
            <th>Path</th>
        </tr>
        $($highRiskServices | ForEach-Object {
            "<tr class='high-risk'>
                <td>$($_.ServerName)</td>
                <td>$($_.ServiceName)</td>
                <td>$($_.StartName)</td>
                <td class='error'>$($_.SecurityFlags)</td>
                <td>$($_.State)</td>
                <td>$($_.PathName)</td>
            </tr>"
        })
    </table>

    <h2>All Services Report</h2>
    <table>
        <tr>
            <th>Server</th>
            <th>Service</th>
            <th>Display Name</th>
            <th>Account</th>
            <th>State</th>
            <th>Start Mode</th>
            <th>Security Flags</th>
            <th>CIS Findings</th>
            <th>Memory (MB)</th>
        </tr>
        $($sortedResults | ForEach-Object {
            $rowClass = if ($_.SecurityFlags -ne "None") { "risk-high" } elseif ($_.CISFindings -ne "Compliant") { "risk-medium" } else { "risk-low" }
            "<tr class='$rowClass'>
                <td>$($_.ServerName)</td>
                <td>$($_.ServiceName)</td>
                <td>$($_.DisplayName)</td>
                <td>$($_.StartName)</td>
                <td>$($_.State)</td>
                <td>$($_.StartMode)</td>
                <td>$($_.SecurityFlags)</td>
                <td>$($_.CISFindings)</td>
                <td>$($_.MemoryUsageMB)</td>
            </tr>"
        })
    </table>
</body>
</html>
"@

    $htmlPath = "$OutputPath\ServiceAccountsReport_$timestamp.html"
    $htmlReport | Out-File -FilePath $htmlPath -Encoding UTF8
    
    return @{
        CSVPath = $csvPath
        HTMLPath = $htmlPath
        Summary = @{
            TotalServers = ($sortedResults | Select-Object -Unique ServerName).Count
            TotalServices = $sortedResults.Count
            HighRiskServices = $highRiskServices.Count
            NonCompliantServices = $nonCompliantServices.Count
            DisabledAccountServices = $disabledAccountServices.Count
        }
    }
}

# Main execution
try {
    Write-Host "Enhanced Service Account Security Scanner" -ForegroundColor Green
    Write-Host "=========================================" -ForegroundColor Green
    Write-Host "Servers to scan: $($servers.Count)"
    Write-Host "AD Lookup enabled: $EnableADLookup"
    Write-Host "Include non-admin services: $IncludeNonAdminServices"
    Write-Host "Max parallel jobs: $MaxParallelJobs"
    Write-Host ""
    
    # Choose scanning method based on server count
    if ($servers.Count -gt 10) {
        Write-Host "Using parallel processing for large server set..." -ForegroundColor Yellow
        $results = Start-ParallelServerScan -Servers $servers -MaxJobs $MaxParallelJobs
    } else {
        Write-Host "Using sequential processing..." -ForegroundColor Yellow
        $results = Start-SequentialServerScan -Servers $servers
    }
    
    if ($results.Count -eq 0) {
        Write-Warning "No services found matching the criteria."
        exit 0
    }
    
    Write-Host "Generating reports..." -ForegroundColor Cyan
    $reportInfo = Export-EnhancedReports -Results $results -OutputPath $OutputPath
    
    # Generate final summary
    $summary = @"

Enhanced Service Account Scan Complete
=====================================
Total Servers Scanned: $($reportInfo.Summary.TotalServers)
Total Services Found: $($reportInfo.Summary.TotalServices)
High Risk Services: $($reportInfo.Summary.HighRiskServices)
Non-Compliant Services: $($reportInfo.Summary.NonCompliantServices)
Services with Disabled Accounts: $($reportInfo.Summary.DisabledAccountServices)

Report Locations:
- CSV Report: $($reportInfo.CSVPath)
- HTML Report: $($reportInfo.HTMLPath)
"@

    Write-Host $summary -ForegroundColor Green
    
    # If there were any errors, notify the user
    if (Test-Path "$OutputPath\ServiceScanErrors.log") {
        Write-Host "`nSome servers could not be scanned. Check $OutputPath\ServiceScanErrors.log for details." -ForegroundColor Yellow
    }
    
    # Open HTML report if available
    if (Test-Path $reportInfo.HTMLPath) {
        $openReport = Read-Host "`nWould you like to open the HTML report? (y/n)"
        if ($openReport -eq 'y' -or $openReport -eq 'Y') {
            Start-Process $reportInfo.HTMLPath
        }
    }
}
catch {
    Write-Error "Script execution failed: $_"
    exit 1
}