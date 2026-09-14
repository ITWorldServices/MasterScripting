$Hostname = $env:COMPUTERNAME

$VMs = (Get-VM -ComputerName $Hostname)

"Host name: $Hostname"
'Number of VMs: ' + ($VMs.Count)
''
$colVMs = @()
foreach ($VM in $VMs)
{
  $objVM = New-Object System.Object
  $objVM | Add-Member -MemberType NoteProperty -Name VMName -Value $VM.VMName
  $objVM | Add-Member -MemberType NoteProperty -Name VMState -Value $VM.State
  $objVM | Add-Member -MemberType NoteProperty -Name VMDynamicMemoryEnabled -Value $VM.DynamicMemoryEnabled
  $objVM | Add-Member -MemberType NoteProperty -Name VMStaticRAM -Value $VM.MemoryStartup 
  if ($vm.DynamicMemoryEnabled) {
    $objVM | Add-Member -MemberType NoteProperty -Name VMDynamicMemoryMax -Value $VM.MemoryMaximum 
  } else {
    $objVM | Add-Member -MemberType NoteProperty -Name VMDynamicMemoryMax -Value 0
  }
  $objVM | Add-Member -MemberType NoteProperty -Name VMCPUCount -Value ([int]$VM.ProcessorCount)
  
  $colVMs += $objVM
}


# display all VMs and their values, nicely formatted
$a = @{Expression={$_.VMName};Label='VM Name'}, `
     @{Expression={$_.VMState};Label='State'}, `
     @{Expression={$_.VMDynamicMemoryEnabled};Label='DynMem enabled'}, `
     @{Expression={('{0:N1}' -f($_.VMStaticRAM/1GB))};Label='Static/startup RAM (GB)';align='right'}, `
     @{Expression={('{0:N1}' -f([Double]$_.VMDynamicMemoryMax/1GB))};Label='Dynamic Mem max (GB)';align='right'}, `
     @{Expression={$_.VMCPUCount};Label='vCPU count';align='right'}

$colVMs | Sort-Object VMName | Format-Table $a -AutoSize

# display sums, max/min, and averages
$b = @{Expression={$_.Property};Label='Property'}, `
     @{Expression={$_.Count};Label='Count'}, `
     @{Expression={('{0:N1}' -f($_.Sum/1GB))};Label='Sum';align='right'}, `
     @{Expression={('{0:N1}' -f($_.Minimum/1GB))};Label='Minimum';align='right'}, `
     @{Expression={('{0:N1}' -f($_.Maximum/1GB))};Label='Maximum';align='right'}, `
     @{Expression={('{0:N1}' -f($_.Average/1GB))};Label='Average';align='right'}

'All VMs'
'======='
$colVMs | Measure-Object VMStaticRAM,VMDynamicMemoryMax -Minimum -Maximum -Sum -Average | Format-Table $b -AutoSize
$colVMs | Measure-Object VMCPUCount -Minimum -Maximum -Sum -Average | Format-Table Property,Count,Sum,Minimum,Maximum,Average -AutoSize

'Running VMs'
'==========='
$colVMs | Where-Object {$_.VMState -eq 'Running'} | Measure-Object VMStaticRAM,VMDynamicMemoryMax -Minimum -Maximum -Sum -Average | Format-Table $b -AutoSize
$colVMs | Where-Object {$_.VMState -eq 'Running'} | Measure-Object VMCPUCount -Minimum -Maximum -Sum -Average | Format-Table Property,Count,Sum,Minimum,Maximum,Average -AutoSize


# add host hardware values
$HostHW = Get-WmiObject Win32_ComputerSystem -ComputerName $Hostname
"Host $Hostname total RAM: " + ('{0:N0}' -f($hostHW.TotalPhysicalMemory / 1GB)  + ' GB')
"Host $Hostname Logical Processors: " + $HostHW.NumberOfLogicalProcessors
