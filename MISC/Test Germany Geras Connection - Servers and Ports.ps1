$hostnames = @(
    'z1sa6803.hdires.zz',
    'z1sa6837.hdires.zz',
    'z1sa0302.hdires.zz',
    'z1sa0303.hdires.zz'
)
$ports = @(1521, 1523, 2484)
$retryIntervalSeconds = 30

while ($true) {
    foreach ($hostname in $hostnames) {
        foreach ($port in $ports) {
            $result = Test-NetConnection $hostname -Port $port -WarningAction SilentlyContinue
            if ($result.TcpTestSucceeded) {
                Write-Host "Connection to $($hostname) on port $($port) is successful."
            } else {
                Write-Host "Connection to $($hostname) on port $($port) failed. Retrying in $retryIntervalSeconds seconds..."
            }

            # Perform a single ping test
            $pingResult = Test-Connection -ComputerName $hostname -Count 1 -ErrorAction SilentlyContinue
            if ($pingResult) {
                Write-Host "Ping to $($hostname) successful. Latency: $($pingResult.ResponseTime) ms"
            } else {
                Write-Host "Ping to $($hostname) failed."
            }
        }
    }
    
    Start-Sleep -Seconds $retryIntervalSeconds
}
