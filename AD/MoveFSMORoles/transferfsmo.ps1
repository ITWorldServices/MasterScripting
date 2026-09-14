#revision0.1 of Kelvin's fsmo move script. it is supe rsimple now. will add checks later
 write-host "Welcome to the Party!  Kelivn created this awesome 5 liner FSMO move script for you"
 $Server = Read-Host -Prompt 'Input your server  name to be the new FSMO holder'
 Move-ADDirectoryServerOperationMasterRole -Identity $server domainnamingmaster
 Move-ADDirectoryServerOperationMasterRole -Identity $Server pdc
 Move-ADDirectoryServerOperationMasterRole -Identity $Server schemamaster
 Move-ADDirectoryServerOperationMasterRole -Identity $Server infrastructuremaster
 Move-ADDirectoryServerOperationMasterRole -Identity $Server rid
#option to run command to check transfer
#netdom /query fsmo