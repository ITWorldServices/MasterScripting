# Import the CSV file from C:\temp\
$serverList = Import-Csv -Path '.\serverlist.csv'

# An array to hold the server info
$serverInfo = @()

# Iterate over each server in the list
foreach ($server in $serverList) {
    Write-Host "Collecting information for $($server.servername)..."

    # Get system info
    $osInfo = Invoke-Command -ComputerName $server.servername {Get-CimInstance -ClassName Win32_OperatingSystem}
    $cpuInfo = Invoke-Command -ComputerName $server.servername {Get-CimInstance -ClassName Win32_Processor}
    $ramInfo = Invoke-Command -ComputerName $server.servername {Get-CimInstance -ClassName Win32_ComputerSystem}
    $netInfo = Invoke-Command -ComputerName $server.servername {Get-NetIPAddress -AddressFamily IPv4}
    $dnsInfo = Invoke-Command -ComputerName $server.servername {Get-DnsClientServerAddress -AddressFamily IPv4}
    $routeInfo = Invoke-Command -ComputerName $server.servername {Get-NetRoute -AddressFamily IPv4 | Where-Object { $_.DestinationPrefix -eq '0.0.0.0/0' }}

    # Count the number of cores and sockets
    $numCores = ($cpuInfo | Measure-Object -Property NumberOfCores -Sum).Sum
    $numSockets = ($cpuInfo | Select-Object -Property SocketDesignation -Unique | Measure-Object).Count

    # Prepare object for data gathering
    $object = [ordered]@{
        'Server Name' = $server.servername
        'OS Name' = $osInfo.Caption
        'OS Version' = $osInfo.Version
        'CPU Cores' = $numCores
        'CPU Sockets' = $numSockets
        'RAM' = $ramInfo.TotalPhysicalMemory / 1GB -as [int]  # RAM in GB
    }

    # Get network details
    for ($i=0; $i -lt $netInfo.Count; $i++) {
        $adapter = $netInfo[$i]
        $adapterDns = $dnsInfo | Where-Object { $_.InterfaceIndex -eq $adapter.InterfaceIndex }
        $adapterGateway = $routeInfo | Where-Object { $_.InterfaceIndex -eq $adapter.InterfaceIndex }

        $object += @{
            "Adapter $(($i+1)) IP Address" = $adapter.IPAddress
            "Adapter $(($i+1)) Subnet" = $adapter.PrefixLength
            "Adapter $(($i+1)) Gateway" = $adapterGateway.NextHop
            "Adapter $(($i+1)) DHCP Server" = $adapter.DhcpServer
            "Adapter $(($i+1)) Primary DNS" = $adapterDns.ServerAddresses[0]
            "Adapter $(($i+1)) Secondary DNS" = $adapterDns.ServerAddresses[1]
            "Adapter $(($i+1)) Tertiary DNS" = $adapterDns.ServerAddresses[2]
        }
    }

    # Add server info to the list
    $serverInfo += New-Object psobject -Property $object

    Write-Host "Information for $($server.servername) collected successfully."
}

Write-Host "Exporting information to Excel file..."

# Export to a new Excel file in C:\temp\ with 'Blue, Table Style Medium 2' format
$serverInfo | Export-Excel -Path 'C:\temp\ServerInfo.xlsx' -TableName 'Table1' -TableStyle 'Medium2' -Show

Write-Host "Export completed successfully."
