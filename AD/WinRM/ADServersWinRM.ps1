Import-Module ActiveDirectory

# Function to test WinRM connectivity and configuration
Function Test-WinRM {
    param (
        [string]$ComputerName
    )

    $result = [PSCustomObject]@{
        ComputerName = $ComputerName
        WinRMService = "Unknown"
        WinRMListener = "Unknown"
        FirewallRule = "Unknown"
        ErrorMessage = ""
    }

    try {
        # Test if the WinRM service is running
        $serviceStatus = Get-Service -Name WinRM -ComputerName $ComputerName -ErrorAction Stop
        $result.WinRMService = if ($serviceStatus.Status -eq 'Running') { "Running" } else { "Stopped" }

        # Check if a WinRM listener is configured
        $listener = Invoke-Command -ComputerName $ComputerName -ScriptBlock {
            winrm enumerate winrm/config/listener
        } -ErrorAction Stop
        $result.WinRMListener = if ($listener) { "Configured" } else { "Not Configured" }

        # Check if the firewall allows WinRM traffic
        $firewallRule = Invoke-Command -ComputerName $ComputerName -ScriptBlock {
            Get-NetFirewallRule -DisplayGroup "Windows Remote Management" -ErrorAction SilentlyContinue |
            Where-Object { $_.Enabled -eq $true -and $_.Direction -eq "Inbound" }
        } -ErrorAction Stop

        $result.FirewallRule = if ($firewallRule) { "Enabled" } else { "Disabled" }
    }
    catch {
        $result.ErrorMessage = $_.Exception.Message
    }

    return $result
}

# Get all AD servers (filter by OperatingSystem containing 'Server')
$ServerComputers = Get-ADComputer -Filter {OperatingSystem -like '*Server*'} -Property Name | Select-Object -ExpandProperty Name

# Output results to a file
$OutputFile = "WinRM_Server_Check_Report.csv"
$Results = @()

foreach ($Computer in $ServerComputers) {
    Write-Host "Checking WinRM configuration on $Computer..." -ForegroundColor Yellow
    $Results += Test-WinRM -ComputerName $Computer
}

# Export results to a CSV file
$Results | Export-Csv -Path $OutputFile -NoTypeInformation -Encoding UTF8

Write-Host "WinRM check completed for servers. Report saved to $OutputFile" -ForegroundColor Green