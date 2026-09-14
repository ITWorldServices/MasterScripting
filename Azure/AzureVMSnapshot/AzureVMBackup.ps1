#############################################
#Backup Metal-Era Azure VM
#v1.0
#Luis Garcia
#Impact Networking
#############################################
#
#Install "Az" PowerShell module version 7.3.0
Install-Module -Name Az -RequiredVersion 7.3.0 -Force
$tenantid = "TENANT-ID"
$subscriptionid = "AZURE-SUBSCRIPTION-ID"
#Connect to Azure
Connect-AzAccount -Tenant $tenantid -SubscriptionId $subscriptionid
#Set Azure Subscription
Set-AzContext -Subscription $subscriptionid
#Connect to Backup Vault
$targetVault = Get-AzRecoveryServicesVault -ResourceGroupName "RESOURCE GROUP NAME" -Name "STORAGE VAULT NAME"
Set-AzRecoveryServicesVaultContext -Vault $targetVault
#Display and input VM name to back up
Get-AzRecoveryServicesBackupContainer -ContainerType "AzureVM"
$friendlyname=Read-Host "Please enter the FriendlyName of the VM you wish to back up"
#Set backup item
$namedContainer = Get-AzRecoveryServicesBackupContainer -ContainerType "AzureVM" -Status "Registered" -FriendlyName $friendlyname -VaultId $targetVault.ID
$item = Get-AzRecoveryServicesBackupItem -Container $namedContainer -WorkloadType "AzureVM" -VaultId $targetVault.ID
#Back up VM
#$endDate = (Get-Date).AddDays(60).ToUniversalTime() #This plus the commented out below would make the backup available for 60 days
$job = Backup-AzRecoveryServicesBackupItem -Item $item -VaultId $targetVault.ID # -ExpiryDateTimeUTC $endDate
#Display progress of backup
$job
Wait-AzRecoveryServicesBackupJob -Job $job -Timeout 43200
$job = Get-AzRecoveryServicesBackupJob -Job $job -VaultId $targetVault.ID
$details = Get-AzRecoveryServicesBackupJobDetail -Job $job -VaultId $targetVault.ID
$details