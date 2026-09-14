# Define list of servers (you can also import from a file if needed)
$servers = Get-Content -Path "servers.txt"


# Create an empty array to store the results
$results = @()

# Loop through each server
foreach ($server in $servers) {
    Write-Host "Scanning $server..." -ForegroundColor Cyan
    try {
        # Get the list of services from the remote server
        $services = Get-WmiObject -Class Win32_Service -ComputerName $server -ErrorAction Stop

        foreach ($service in $services) {
            # Store service details in a custom object
            $results += [PSCustomObject]@{
                ServerName     = $server
                ServiceName    = $service.Name
                DisplayName    = $service.DisplayName
                StartMode      = $service.StartMode
                State          = $service.State
                StartName      = $service.StartName  # This is the account running the service
            }
        }
    } catch {
        Write-Warning "Failed to connect to $server: $_"
    }
}

# Export the results to a CSV file
$csvPath = "ServiceAccountsReport.csv"
$results | Export-Csv -Path $csvPath -NoTypeInformation -Encoding UTF8

Write-Host "Scan complete. Results saved to $csvPath" -ForegroundColor Green
