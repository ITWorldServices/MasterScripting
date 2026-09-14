#Requires -Version 5.1
<#
.SYNOPSIS
Collects basic inventory from legacy Windows Servers using WMI/DCOM.

.DESCRIPTION
This function is dot-sourced by Invoke-ADServerDiscovery.ps1. It is intended
for Windows Server 2003 and 2008 systems that cannot execute the modern remote
collector. It does not copy files to or run PowerShell on the target server.
#>
function Get-LegacyServerDiscoveryData {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$ComputerName,
        [pscredential]$Credential
    )

    $diagnostics=New-Object 'System.Collections.Generic.List[object]'
    function Add-LegacyDiagnostic {
        param([string]$Collector,[string]$Status,[string]$Message,[int]$Count=0)
        $diagnostics.Add([pscustomobject][ordered]@{
            Time=Get-Date
            ComputerName=$ComputerName
            Collector=$Collector
            Status=$Status
            RecordCount=$Count
            Message=$Message
        })
    }
    function Invoke-LegacyWmi {
        param([string]$Class,[string]$Namespace='root\cimv2',[string]$Filter,[switch]$List)
        $parameters=@{ComputerName=$ComputerName;Namespace=$Namespace;Class=$Class;ErrorAction='Stop'}
        if($Filter){$parameters.Filter=$Filter}
        if($List){$parameters.List=$true}
        if($Credential){$parameters.Credential=$Credential}
        Get-WmiObject @parameters
    }
    function Invoke-LegacyCollector {
        param([string]$Name,[scriptblock]$Action)
        try{
            $result=@(& $Action)
            Add-LegacyDiagnostic $Name Success 'Legacy collector completed.' $result.Count
            Write-Output -NoEnumerate $result
        }catch{
            Add-LegacyDiagnostic $Name Failed $_.Exception.Message
            Write-Output -NoEnumerate @()
        }
    }
    function Join-LegacyValue {
        param([object[]]$Values,[string]$Separator='; ')
        $joined=(($Values | Where-Object {$null -ne $_ -and -not [string]::IsNullOrWhiteSpace([string]$_)} | ForEach-Object {([string]$_).Trim()} | Sort-Object -Unique) -join $Separator)
        if([string]::IsNullOrWhiteSpace($joined)){return $null}
        $joined
    }
    function Convert-LegacyDate {
        param($Value)
        if(-not $Value){return $null}
        try{[Management.ManagementDateTimeConverter]::ToDateTime([string]$Value)}catch{$null}
    }
    function Get-WmiReferencePart {
        param([string]$Reference,[string]$Name)
        $pattern=('{0}="([^"]*)"' -f [regex]::Escape($Name))
        if($Reference -match $pattern){return $matches[1].Replace('\\','\')}
        return $null
    }

    $cs=$null
    $system=Invoke-LegacyCollector 'System' {
        $cs=Invoke-LegacyWmi Win32_ComputerSystem | Select-Object -First 1
        $os=Invoke-LegacyWmi Win32_OperatingSystem | Select-Object -First 1
        $cpu=@(Invoke-LegacyWmi Win32_Processor)
        $bios=Invoke-LegacyWmi Win32_BIOS | Select-Object -First 1
        $virtual=(
            $cs.Manufacturer -match 'VMware|QEMU|Xen|innotek' -or
            $cs.Model -match 'Virtual Machine|VirtualBox|KVM|HVM domU|VMware|Bochs|Xen'
        )
        [pscustomobject][ordered]@{
            ComputerName=$cs.Name
            Domain=$cs.Domain
            Manufacturer=$cs.Manufacturer
            Model=$cs.Model
            SerialNumber=$bios.SerialNumber
            PhysicalOrVirtual=$(if($virtual){'Virtual'}else{'Physical'})
            OperatingSystem=$os.Caption
            Version=$os.Version
            Build=$os.BuildNumber
            LastBoot=(Convert-LegacyDate $os.LastBootUpTime)
            CPU=(Join-LegacyValue $cpu.Name)
            Sockets=$cpu.Count
            Cores=($cpu | Measure-Object NumberOfCores -Sum).Sum
            LogicalProcessors=($cpu | Measure-Object NumberOfLogicalProcessors -Sum).Sum
            MemoryGB=[math]::Round([double]$cs.TotalPhysicalMemory/1GB,2)
        }
    }

    if(-not $system.Count){
        throw ('Legacy WMI system inventory failed for {0}: {1}' -f $ComputerName,(Join-LegacyValue @($diagnostics | Where-Object Status -eq Failed | ForEach-Object Message)))
    }
    $cs=Invoke-LegacyWmi Win32_ComputerSystem | Select-Object -First 1

    $network=Invoke-LegacyCollector 'Network' {
        foreach($adapter in @(Invoke-LegacyWmi Win32_NetworkAdapterConfiguration -Filter 'IPEnabled=True')){
            $addresses=@($adapter.IPAddress)
            for($index=0;$index -lt $addresses.Count;$index++){
                $address=[string]$addresses[$index]
                if($address -notmatch '^\d{1,3}(\.\d{1,3}){3}$'){continue}
                [pscustomobject][ordered]@{
                    ComputerName=$ComputerName
                    Interface=$adapter.Description
                    IPAddress=$address
                    PrefixLength=$null
                    SubnetMask=$(if(@($adapter.IPSubnet).Count -gt $index){@($adapter.IPSubnet)[$index]}else{$null})
                    Gateway=(Join-LegacyValue @($adapter.DefaultIPGateway))
                    DNS=(Join-LegacyValue @($adapter.DNSServerSearchOrder))
                    MacAddress=$adapter.MACAddress
                }
            }
        }
    }

    $storage=Invoke-LegacyCollector 'Storage' {
        Invoke-LegacyWmi Win32_LogicalDisk -Filter 'DriveType=3' | Sort-Object DeviceID | ForEach-Object {
            [pscustomobject][ordered]@{
                ComputerName=$ComputerName
                Drive=$_.DeviceID
                Label=$_.VolumeName
                FileSystem=$_.FileSystem
                SizeGB=[math]::Round([double]$_.Size/1GB,2)
                FreeGB=[math]::Round([double]$_.FreeSpace/1GB,2)
                PercentFree=$(if($_.Size){[math]::Round(100*[double]$_.FreeSpace/[double]$_.Size,2)}else{$null})
            }
        }
    }

    $rolesAndFeatures=Invoke-LegacyCollector 'Roles and Features' {
        Invoke-LegacyWmi Win32_ServerFeature | Sort-Object Name | ForEach-Object {
            [pscustomobject][ordered]@{
                ComputerName=$ComputerName
                Name=$_.Name
                DisplayName=$_.Name
                FeatureType='Legacy server feature'
            }
        }
    }

    $applications=Invoke-LegacyCollector 'Applications' {
        $registry=Invoke-LegacyWmi StdRegProv 'root\default' -List
        $hklm=2147483650
        $basePaths='SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall','SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall'
        foreach($basePath in $basePaths){
            foreach($subKey in @($registry.EnumKey($hklm,$basePath).sNames)){
                $path='{0}\{1}' -f $basePath,$subKey
                $name=$registry.GetStringValue($hklm,$path,'DisplayName').sValue
                if([string]::IsNullOrWhiteSpace($name)){continue}
                [pscustomobject][ordered]@{
                    ComputerName=$ComputerName
                    Name=$name
                    Version=$registry.GetStringValue($hklm,$path,'DisplayVersion').sValue
                    Publisher=$registry.GetStringValue($hklm,$path,'Publisher').sValue
                    InstallDate=$registry.GetStringValue($hklm,$path,'InstallDate').sValue
                    InstallLocation=$registry.GetStringValue($hklm,$path,'InstallLocation').sValue
                }
            }
        }
    }

    $services=Invoke-LegacyCollector 'Services' {
        Invoke-LegacyWmi Win32_Service | Sort-Object DisplayName | ForEach-Object {
            [pscustomobject][ordered]@{
                ComputerName=$ComputerName
                Name=$_.Name
                DisplayName=$_.DisplayName
                State=$_.State
                StartMode=$_.StartMode
                StartName=$_.StartName
                PathName=$_.PathName
            }
        }
    }

    $isSqlServer=@($services | Where-Object {$_.Name -match '^(MSSQLSERVER|MSSQL\$|SQLSERVERAGENT|SQLAgent\$|MSOLAP\$|MSSQLServerOLAPService|ReportServer|ReportServer\$)'}).Count -gt 0
    $isExchangeServer=@($services | Where-Object {$_.Name -match '^MSExchange'}).Count -gt 0
    $serverClass=if($isSqlServer -and $isExchangeServer){'SQL and Exchange'}elseif($isSqlServer){'SQL Server'}elseif($isExchangeServer){'Exchange Server'}else{'General Server'}
    foreach($row in $system){
        $row | Add-Member NoteProperty ServerClass $serverClass -Force
        $row | Add-Member NoteProperty CollectionMode 'Legacy basic only' -Force
    }

    $listeningPorts=@()
    Add-LegacyDiagnostic 'Listening Ports' Skipped 'The legacy WMI/DCOM path cannot reliably query listening ports.'

    $shares=Invoke-LegacyCollector 'Shares and Permissions' {
        Invoke-LegacyWmi Win32_Share | Sort-Object Name | ForEach-Object {
            [pscustomobject][ordered]@{
                ComputerName=$ComputerName
                Name=$_.Name
                Path=$_.Path
                UNCPath=('\\{0}\{1}' -f $ComputerName,$_.Name)
                Description=$_.Description
                Special=([int64]$_.Type -ge 2147483648)
                SharePermissions=$null
                NTFSPermissions=$null
            }
        }
    }
    Add-LegacyDiagnostic 'Share Permissions' Warning 'Legacy WMI inventory returned shares without expanded share or NTFS ACLs.' $shares.Count

    $scheduledTasks=Invoke-LegacyCollector 'Scheduled Tasks' {
        $raw=@(& schtasks.exe /Query /S $ComputerName /FO CSV /V 2>&1)
        if($LASTEXITCODE -ne 0){throw (Join-LegacyValue $raw ' ')}
        $raw | ConvertFrom-Csv | ForEach-Object {
            [pscustomobject][ordered]@{
                ComputerName=$ComputerName
                Name=$_.TaskName
                TaskPath=$_.TaskName
                Executable=$_.'Task To Run'
                State=$_.Status
                RunAsUser=$_.'Run As User'
                Description=$_.Comment
            }
        }
    }

    $localAccounts=Invoke-LegacyCollector 'Local Accounts' {
        Invoke-LegacyWmi Win32_UserAccount -Filter 'LocalAccount=True' | Sort-Object Name | ForEach-Object {
            [pscustomobject][ordered]@{
                ComputerName=$ComputerName
                Name=$_.Name
                Enabled=(-not $_.Disabled)
                Description=$_.Description
                LastLogon=$null
                PasswordExpires=$_.PasswordExpires
            }
        }
    }

    $localGroupMembership=Invoke-LegacyCollector 'Local Group Membership' {
        Invoke-LegacyWmi Win32_GroupUser | ForEach-Object {
            $groupDomain=Get-WmiReferencePart $_.GroupComponent 'Domain'
            $groupName=Get-WmiReferencePart $_.GroupComponent 'Name'
            $memberDomain=Get-WmiReferencePart $_.PartComponent 'Domain'
            $memberName=Get-WmiReferencePart $_.PartComponent 'Name'
            [pscustomobject][ordered]@{
                ComputerName=$ComputerName
                Group=$(if($groupDomain -and $groupDomain -ne $ComputerName){'{0}\{1}' -f $groupDomain,$groupName}else{$groupName})
                Name=$(if($memberDomain){'{0}\{1}' -f $memberDomain,$memberName}else{$memberName})
                ObjectClass=$(if($_.PartComponent -match 'Win32_UserAccount'){'User'}elseif($_.PartComponent -match 'Win32_Group'){'Group'}else{'Unknown'})
                PrincipalSource='WMI'
            }
        }
    }

    $serviceAccounts=@($services | Where-Object {
        $_.StartName -and $_.StartName -notmatch '^(LocalSystem|LocalService|NetworkService|NT AUTHORITY\\|NT SERVICE\\)'
    } | ForEach-Object {
        [pscustomobject][ordered]@{
            ComputerName=$ComputerName
            Account=$_.StartName
            Usage=('Windows service: {0}' -f $_.DisplayName)
            InteractiveLoginRequired='Unknown'
            ManagedServiceAccount=$(if($_.StartName.EndsWith('$')){'Yes'}else{'No'})
        }
    } | Sort-Object Account,Usage -Unique)
    Add-LegacyDiagnostic 'Service Accounts' Success 'Legacy collector completed.' $serviceAccounts.Count

    foreach($collectorName in 'Active Directory','AD Users','AD Group Membership','AD Replication','DHCP Scopes','DHCP Exclusions','DHCP Reservations','DHCP Options','DHCP Failover','DNS Zones'){
        Add-LegacyDiagnostic $collectorName Skipped 'Legacy basic-only collection does not run role-specific collectors.'
    }

    $hasActiveDirectory=$false
    if($cs -and [int]$cs.DomainRole -ge 4){$hasActiveDirectory=$true}
    $featureNames=Join-LegacyValue @($rolesAndFeatures.Name)
    [pscustomobject][ordered]@{
        Metadata=[pscustomobject][ordered]@{
            ComputerName=$system[0].ComputerName
            ServerClass=$serverClass
            CollectionMode='Legacy basic only'
            IsSqlServer=$isSqlServer
            IsExchangeServer=$isExchangeServer
            HasActiveDirectory=$hasActiveDirectory
            HasDhcp=($featureNames -match 'DHCP')
            HasDns=($featureNames -match 'DNS')
            CollectedAt=Get-Date
        }
        System=@($system)
        Network=@($network)
        Storage=@($storage)
        RolesAndFeatures=@($rolesAndFeatures)
        Applications=@($applications)
        Services=@($services)
        ListeningPorts=@($listeningPorts)
        SharesAndPermissions=@($shares)
        ScheduledTasks=@($scheduledTasks)
        LocalAccounts=@($localAccounts)
        LocalGroupMembership=@($localGroupMembership)
        ServiceAccounts=@($serviceAccounts)
        ActiveDirectory=@()
        ADUsers=@()
        ADGroupMembership=@()
        ADReplication=@()
        DHCPScopes=@()
        DHCPExclusions=@()
        DHCPReservations=@()
        DHCPOptions=@()
        DHCPFailover=@()
        DNSZones=@()
        Diagnostics=@($diagnostics)
    }
}
