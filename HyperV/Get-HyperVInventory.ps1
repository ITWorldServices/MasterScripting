# Requires: Imported HyperV PowerShell module (http://pshyperv.codeplex.com/releases/view/62842)
# Requires: Import-Module FailoverClusters
# Requires: Running PowerShell as Administrator in order to properly import the above modules
 
function Get-HyperVInventory {
<#
.SYNOPSIS
Fetches Hyper-V VM Inventory from a specified Hyper-V Failover cluster
 
.DESCRIPTION
Fetches Hyper-V VM Inventory from a specified Hyper-V Failover cluster
 
.PARAMETER ClusterName
The Name of the Hyper-V Failover Cluster to inspect
 
.EXAMPLE
PS F:\> Get-HyperVInventory -ClusterName "dev-cluster1"
 
.EXAMPLE
PS F:\> Get-Cluster | Get-HyperVInventory
 
.LINK
http://www.shogan.co.uk
 
.NOTES
Created by: Sean Duffy
Date: 09/07/2012
#>
 
[CmdletBinding()]
param(
[Parameter(Position=0,Mandatory=$true,HelpMessage="Name of the Cluster to fetch inventory from",
ValueFromPipeline=$true,ValueFromPipelineByPropertyName=$true)]
[System.String]
$ClusterName
)
 
process {
 
$Report = @()
 
$Cluster = Get-Cluster -Name $ClusterName
$HVHosts = $Cluster | Get-ClusterNode
 
foreach ($HVHost in $HVHosts) {
$VMs = Get-VM -Server $HVHost
foreach ($VM in $VMs) {
[long]$TotalVHDSize = 0
$VHDCount = 0
$VMName = $VM.VMElementName
$VMMemory = $VM | Get-VMMemory
$CPUCount = $VM | Get-VMCPUCount
$NetSwitch = $VM | Get-VMNIC
$NetMacAdd = $VM | Get-VMNIC
# VM Disk Info
$VHDDisks = $VM | Get-VMDisk | Where { $_.DiskName -like "Hard Disk Image" }
foreach ($disk in $VHDDisks) {
$VHDInfo = Get-VHDInfo -VHDPaths $disk.DiskImage
$TotalVHDSize = $TotalVHDSize + $VHDInfo.FileSize
$VHDCount += 1
}
$TotalVHDSize = $TotalVHDSize/1024/1024
$row = New-Object -Type PSObject -Property @{
Cluster = $Cluster.Name
VMName = $VMName
VMMemory = $VMMemory.VirtualQuantity
CPUCount = $CPUCount.VirtualQuantity
CPUSocketCount = $CPUCount.SocketCount
NetSwitch = $NetSwitch.SwitchName
NetMACAdd = $NetMacAdd.Address
HostName = $HVHost.Name
VMState = $HVHost.State
TotalVMDiskSizeMB = $TotalVHDSize
TotalVMDiskCount = $VHDCount
} ## end New-Object
$Report += $row
}
}
return $Report
 
}
}