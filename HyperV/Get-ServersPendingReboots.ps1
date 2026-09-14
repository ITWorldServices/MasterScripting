Import-Module ActiveDirectory
$servers = Get-ADComputer -Filter * -Properties operatingsystem |
    Where operatingsystem -match 'server'
$path = 'HKLM:\SYSTEM\CurrentControlSet\Control\Session Manager'
$name = 'PendingFileRenameOperations'
Invoke-command -ComputerName $servers.name -Ea 0 -ScriptBlock {
  Get-ItemProperty -Path $using:path -Name $using:name}   
Select-Object pscomputername, @{
 LABEL='RebootRequired';
 EXPRESSION={if($_.PendingFileRenameOperations){$true}}}
Export-CSV -Path C:\scripts\PendingReboots.csv




