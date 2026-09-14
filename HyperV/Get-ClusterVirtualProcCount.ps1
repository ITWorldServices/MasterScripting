$Clusters = Get-Content C:\scripts\Cluster.txt
foreach ($Cluster in $Clusters)
{
	[int]$TotalFreeMemory = 0;
	[int]$TotalMemory = 0;
	$ClusterNodes = Get-Cluster $Cluster | Get-ClusterNode
	foreach ($ClusterNode in $ClusterNodes)
	{
		[int]$FreeMemory = [math]::round(((Get-WmiObject -ComputerName $ClusterNode -Class Win32_OperatingSystem).FreePhysicalMemory / 1MB), 0)
		[int]$TotalFreeMemory = [int]$TotalFreeMemory + [int]$FreeMemory
		[int]$NodeMemory = [math]::round(((Get-WmiObject -ComputerName $ClusterNode -Class Win32_OperatingSystem).TotalVisibleMemorySize / 1MB), 0)
		[int]$TotalMemory = [int]$TotalMemory + [int]$NodeMemory
	}
	[int]$TotalAvailableMemory = [int]$TotalFreeMemory - [int]$NodeMemory
	
	Write-Host "Cluster: $Cluster"
	Write-Host "Total Memory: $TotalMemory"
	Write-Host "Total Free Memory: $TotalFreeMemory"
	Write-Host "Total Available Memory: $TotalAvailableMemory"
	Write-Host " "
}