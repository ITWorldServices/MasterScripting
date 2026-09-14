# Set the backup location
$backupLocation = "C:\ExchangeConfigBackup"
$timestamp = Get-Date -Format "yyyyMMddHHmmss"

# Create the backup folder
$backupFolder = Join-Path -Path $backupLocation -ChildPath $timestamp
New-Item -Path $backupFolder -ItemType Directory -Force

# Get the Exchange installation path
$exchangeInstallPath = [System.Environment]::GetEnvironmentVariable("ExchangeInstallPath", "Machine")

# Paths to Exchange configuration files
$configFiles = @(
	"Bin\ComplianceAuditService.exe.config",
	"Bin\EdgeTransport.exe.config",
	"Bin\Microsoft.Exchange.Diagnostics.Service.exe.config",
	"Bin\Microsoft.Exchange.Directory.TopologyService.exe.config",
	"Bin\Microsoft.Exchange.EdgeSyncSvc.exe.config",
	"Bin\Microsoft.Exchange.Mitigation.Service.exe.config",
	"Bin\Microsoft.Exchange.RpcClientAccess.Service.exe.config",
	"Bin\Microsoft.Exchange.Search.Service.exe.config",
	"Bin\Microsoft.Exchange.Servicehost.exe.config",
	"Bin\Microsoft.Exchange.Store.Service.exe.config",
	"Bin\MSExchangeCompliance.exe.config",
	"Bin\MSExchangeDelivery.exe.config",
	"Bin\MSExchangeFrontEndTransport.exe.config",
	"Bin\MSExchangeHMHost.exe.config",
	"Bin\MSExchangeHMRecovery.exe.config",
	"Bin\MSExchangeHMWorker.exe.config",
	"Bin\MSExchangeMailboxAssistants.exe.config",
	"Bin\MsExchangeMailboxReplication.exe.config",
	"Bin\MSExchangeSubmission.exe.config",
	"Bin\MSExchangeThrottling.exe.config",
	"Bin\MSExchangeTransport.exe.config",
	"ClientAccess\PopImap\Microsoft.Exchange.Imap4.exe.config",
	"ClientAccess\PopImap\Microsoft.Exchange.Imap4Service.exe.config",
	"ClientAccess\PopImap\Microsoft.Exchange.Pop3.exe.config",
	"ClientAccess\PopImap\Microsoft.Exchange.Pop3Service.exe.config",
	"FrontEnd\PopImap\Microsoft.Exchange.Imap4.exe.config",
	"FrontEnd\PopImap\Microsoft.Exchange.Imap4Service.exe.config",
	"FrontEnd\PopImap\Microsoft.Exchange.Pop3.exe.config",
	"FrontEnd\PopImap\Microsoft.Exchange.Pop3Service.exe.config",
	"Bin\Microsoft.Exchange.AddressBook.Service.dll.config",
	"Bin\Microsoft.Exchange.Management.Transport.dll.config",
	"TransportRoles\agents\Antimalware\Microsoft.Exchange.Transport.Agent.Malware.dll.config",
	"Bin\MSExchangeUM.config",
	"ClientAccess\Autodiscover\web.config",
	"ClientAccess\ecp\web.config",
	"ClientAccess\ecp\DLPPolicy\Web.config",
	"ClientAccess\ecp\Handlers\Web.config",
	"ClientAccess\ecp\PersonalSettings\Web.config",
	"ClientAccess\ecp\UsersGroups\Web.config",
	"ClientAccess\exchweb\ews\web.config",
	"ClientAccess\mapi\emsmdb\web.config",
	"ClientAccess\mapi\nspi\web.config",
	"ClientAccess\OAB\web.config",
	"ClientAccess\PowerShell\web.config",
	"ClientAccess\PowerShell-Proxy\web.config",
	"ClientAccess\PushNotifications\web.config",
	"ClientAccess\rest\web.config",
	"ClientAccess\RpcProxy\web.config",
	"ClientAccess\Sync\web.config",
	"FrontEnd\HttpProxy\autodiscover\web.config",
	"FrontEnd\HttpProxy\ecp\web.config",
	"FrontEnd\HttpProxy\ews\web.config",
	"FrontEnd\HttpProxy\mapi\web.config",
	"FrontEnd\HttpProxy\oab\web.config",
	"FrontEnd\HttpProxy\owa\web.config",
	"FrontEnd\HttpProxy\powershell\web.config",
	"FrontEnd\HttpProxy\pushnotifications\web.config",
	"FrontEnd\HttpProxy\ReportingWebService\web.config",
	"FrontEnd\HttpProxy\rest\web.config",
	"FrontEnd\HttpProxy\rpc\web.config",
	"FrontEnd\HttpProxy\sync\web.config",
	"Bin\Search\Ceres\Runtime\1.0\Noderunner.exe.config"
)

# Backup configuration files
foreach ($file in $configFiles) {
    $filePath = Join-Path -Path $exchangeInstallPath -ChildPath $file
    $fileName = [System.IO.Path]::GetFileName($file)
    $destinationPath = Join-Path -Path $backupFolder -ChildPath $fileName
    Copy-Item -Path $filePath -Destination $destinationPath -Force
}

Write-Host "Exchange configuration files backed up to: $backupFolder"