#Requires -Version 4.0
<#
.SYNOPSIS
Collects discovery data from the Windows Server where this script executes.

.DESCRIPTION
This data-only collector is designed for local execution or Invoke-Command
-FilePath. It does not install modules or create files. SQL Server and Exchange
Server systems are automatically limited to basic inventory. AD DS, DHCP, and
DNS collectors run only when the matching role is installed and the server is
not classified as SQL or Exchange.
#>
[CmdletBinding()]
param(
    [bool]$IncludeMicrosoftTasks = $false
)

$ErrorActionPreference='Stop'
$ProgressPreference='SilentlyContinue'
$script:Diagnostics=New-Object 'System.Collections.Generic.List[object]'
$ComputerName=$env:COMPUTERNAME

function Add-DiscoveryDiagnostic {
    param(
        [string]$Collector,
        [ValidateSet('Success','Warning','Failed','Skipped')][string]$Status,
        [string]$Message,
        [int]$Count=0
    )
    $script:Diagnostics.Add([pscustomobject][ordered]@{
        Time=(Get-Date)
        ComputerName=$ComputerName
        Collector=$Collector
        Status=$Status
        RecordCount=$Count
        Message=$Message
    })
}

function Invoke-DiscoveryCollector {
    param([string]$Name,[scriptblock]$Action)
    try {
        $result=@(& $Action)
        Add-DiscoveryDiagnostic $Name Success 'Collector completed.' $result.Count
        Write-Output -NoEnumerate $result
    } catch {
        Add-DiscoveryDiagnostic $Name Failed $_.Exception.Message
        Write-Warning ('{0}: {1}' -f $Name,$_.Exception.Message)
        Write-Output -NoEnumerate @()
    }
}

function Join-DiscoveryValue {
    param([object[]]$Values,[string]$Separator='; ')
    $joined=(($Values | Where-Object {
        $null -ne $_ -and -not [string]::IsNullOrWhiteSpace([string]$_)
    } | ForEach-Object {
        ([string]$_).Trim()
    } | Sort-Object -Unique) -join $Separator)
    if([string]::IsNullOrWhiteSpace($joined)){return $null}
    return $joined
}

function Convert-PrefixToMask {
    param([ValidateRange(0,32)][int]$Prefix)
    $bits=(('1' * $Prefix) -join '').PadRight(32,'0')
    return ((0,8,16,24 | ForEach-Object {
        [Convert]::ToInt32($bits.Substring($_,8),2)
    }) -join '.')
}

if($env:OS -ne 'Windows_NT'){
    throw 'This collector must run on Windows.'
}

$system=Invoke-DiscoveryCollector 'System' {
    $cs=Get-CimInstance Win32_ComputerSystem
    $os=Get-CimInstance Win32_OperatingSystem
    $cpu=@(Get-CimInstance Win32_Processor)
    $bios=Get-CimInstance Win32_BIOS
    # HypervisorPresent is also true on physical Hyper-V hosts. Classify from
    # hardware vendor/model signatures so a Dell/HP host is not marked virtual.
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
        LastBoot=$os.LastBootUpTime
        CPU=(Join-DiscoveryValue $cpu.Name)
        Sockets=$cpu.Count
        Cores=($cpu | Measure-Object NumberOfCores -Sum).Sum
        LogicalProcessors=($cpu | Measure-Object NumberOfLogicalProcessors -Sum).Sum
        MemoryGB=[math]::Round($cs.TotalPhysicalMemory/1GB,2)
    }
}

$network=Invoke-DiscoveryCollector 'Network' {
    foreach($config in Get-NetIPConfiguration | Where-Object {$_.NetAdapter.Status -eq 'Up'}){
        foreach($ip in @($config.IPv4Address)){
            [pscustomobject][ordered]@{
                ComputerName=$ComputerName
                Interface=$config.InterfaceAlias
                IPAddress=$ip.IPAddress
                PrefixLength=$ip.PrefixLength
                SubnetMask=(Convert-PrefixToMask $ip.PrefixLength)
                Gateway=(Join-DiscoveryValue $config.IPv4DefaultGateway.NextHop)
                DNS=(Join-DiscoveryValue @($config.DNSServer.ServerAddresses | Where-Object {
                    ($_ -as [ipaddress]) -and
                    ([ipaddress]$_).AddressFamily -eq [Net.Sockets.AddressFamily]::InterNetwork
                }))
                MacAddress=$config.NetAdapter.MacAddress
            }
        }
    }
}

$storage=Invoke-DiscoveryCollector 'Storage' {
    Get-CimInstance Win32_LogicalDisk -Filter 'DriveType=3' |
      Sort-Object DeviceID | ForEach-Object {
        [pscustomobject][ordered]@{
            ComputerName=$ComputerName
            Drive=$_.DeviceID
            Label=$_.VolumeName
            FileSystem=$_.FileSystem
            SizeGB=[math]::Round($_.Size/1GB,2)
            FreeGB=[math]::Round($_.FreeSpace/1GB,2)
            PercentFree=$(if($_.Size){[math]::Round(100*$_.FreeSpace/$_.Size,2)}else{$null})
        }
      }
}

$rolesAndFeatures=Invoke-DiscoveryCollector 'Roles and Features' {
    if(-not(Get-Command Get-WindowsFeature -ErrorAction SilentlyContinue)){
        throw 'Get-WindowsFeature is unavailable.'
    }
    Get-WindowsFeature | Where-Object Installed | ForEach-Object {
        [pscustomobject][ordered]@{
            ComputerName=$ComputerName
            Name=$_.Name
            DisplayName=$_.DisplayName
            FeatureType=$_.FeatureType
        }
    }
}

$applications=Invoke-DiscoveryCollector 'Applications' {
    $paths=@(
        'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*',
        'HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*'
    )
    Get-ItemProperty $paths -ErrorAction SilentlyContinue |
      Where-Object {$_.DisplayName -and $_.SystemComponent -ne 1} |
      Sort-Object DisplayName,DisplayVersion -Unique | ForEach-Object {
        [pscustomobject][ordered]@{
            ComputerName=$ComputerName
            Name=$_.DisplayName
            Version=$_.DisplayVersion
            Publisher=$_.Publisher
            InstallDate=$_.InstallDate
            InstallLocation=$_.InstallLocation
        }
      }
}

$services=Invoke-DiscoveryCollector 'Services' {
    Get-CimInstance Win32_Service | Sort-Object DisplayName | ForEach-Object {
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

$isSqlServer=@($services | Where-Object {
    $_.Name -match '^(MSSQLSERVER|MSSQL\$|SQLSERVERAGENT|SQLAgent\$|MSOLAP\$|MSSQLServerOLAPService|ReportServer|ReportServer\$)'
}).Count -gt 0
$isExchangeServer=@($services | Where-Object {$_.Name -match '^MSExchange'}).Count -gt 0
$serverClass=if($isSqlServer -and $isExchangeServer){
    'SQL and Exchange'
}elseif($isSqlServer){
    'SQL Server'
}elseif($isExchangeServer){
    'Exchange Server'
}else{
    'General Server'
}
$basicOnly=$isSqlServer -or $isExchangeServer
$collectionMode=if($basicOnly){'Basic only'}else{'Full role-aware'}

foreach($row in $system){
    $row | Add-Member NoteProperty ServerClass $serverClass -Force
    $row | Add-Member NoteProperty CollectionMode $collectionMode -Force
}

$listeningPorts=Invoke-DiscoveryCollector 'Listening Ports' {
    Get-NetTCPConnection -State Listen |
      Sort-Object LocalPort,LocalAddress -Unique | ForEach-Object {
        [pscustomobject][ordered]@{
            ComputerName=$ComputerName
            LocalAddress=$_.LocalAddress
            LocalPort=$_.LocalPort
            ProcessId=$_.OwningProcess
            ProcessName=$(try{
                (Get-Process -Id $_.OwningProcess -ErrorAction Stop).ProcessName
            }catch{$null})
        }
      }
}

$shares=Invoke-DiscoveryCollector 'Shares and Permissions' {
    if(-not(Get-Command Get-SmbShare -ErrorAction SilentlyContinue)){
        throw 'Get-SmbShare is unavailable.'
    }
    Get-SmbShare | Sort-Object Name | ForEach-Object {
        $share=$_
        $shareAcl=try{
            Get-SmbShareAccess $share.Name | ForEach-Object {
                '{0}: {1} ({2})' -f $_.AccountName,$_.AccessRight,$_.AccessControlType
            }
        }catch{@()}
        $ntfsAcl=try{
            if($share.Path -and (Test-Path -LiteralPath $share.Path)){
                (Get-Acl -LiteralPath $share.Path).Access | ForEach-Object {
                    '{0}: {1} ({2})' -f $_.IdentityReference,$_.FileSystemRights,$_.AccessControlType
                }
            }
        }catch{@()}
        [pscustomobject][ordered]@{
            ComputerName=$ComputerName
            Name=$share.Name
            Path=$share.Path
            UNCPath=('\\{0}\{1}' -f $ComputerName,$share.Name)
            Description=$share.Description
            Special=$share.Special
            SharePermissions=(Join-DiscoveryValue $shareAcl)
            NTFSPermissions=(Join-DiscoveryValue $ntfsAcl)
        }
    }
}

$scheduledTasks=Invoke-DiscoveryCollector 'Scheduled Tasks' {
    if(-not(Get-Command Get-ScheduledTask -ErrorAction SilentlyContinue)){
        throw 'Get-ScheduledTask is unavailable.'
    }
    Get-ScheduledTask | Where-Object {
        $IncludeMicrosoftTasks -or $_.TaskPath -notlike '\Microsoft\*'
    } | Sort-Object TaskPath,TaskName | ForEach-Object {
        [pscustomobject][ordered]@{
            ComputerName=$ComputerName
            Name=$_.TaskName
            TaskPath=$_.TaskPath
            Executable=(Join-DiscoveryValue ($_.Actions | ForEach-Object {
                ('{0} {1}' -f $_.Execute,$_.Arguments).Trim()
            }))
            State=$_.State
            RunAsUser=$_.Principal.UserId
            Description=$_.Description
        }
    }
}

$localAccounts=Invoke-DiscoveryCollector 'Local Accounts' {
    if($rolesAndFeatures.Name -contains 'AD-Domain-Services'){return}
    if(Get-Command Get-LocalUser -ErrorAction SilentlyContinue){
        Get-LocalUser | Sort-Object Name | Select-Object @{n='ComputerName';e={$ComputerName}},Name,Enabled,Description,LastLogon,PasswordExpires
        return
    }
    Add-DiscoveryDiagnostic 'Local Accounts Compatibility' Warning 'Get-LocalUser is unavailable; using Win32_UserAccount.'
    Get-CimInstance Win32_UserAccount -Filter 'LocalAccount=True' | Sort-Object Name | ForEach-Object {
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

$localGroupMembership=Invoke-DiscoveryCollector 'Local Group Membership' {
    if($rolesAndFeatures.Name -contains 'AD-Domain-Services'){return}
    if(Get-Command Get-LocalGroupMember -ErrorAction SilentlyContinue){
        foreach($group in Get-LocalGroup){
            try{
                Get-LocalGroupMember $group.Name -ErrorAction Stop | Select-Object @{n='ComputerName';e={$ComputerName}},@{n='Group';e={$group.Name}},Name,ObjectClass,PrincipalSource
            }catch{
                Add-DiscoveryDiagnostic 'Local Group Membership' Warning ('{0}: {1}' -f $group.Name,$_.Exception.Message)
            }
        }
        return
    }
    Add-DiscoveryDiagnostic 'Local Group Membership Compatibility' Warning 'Local group cmdlets are unavailable; using the WinNT provider.'
    $computer=[ADSI]('WinNT://{0},computer' -f $ComputerName)
    foreach($group in @($computer.psbase.Children | Where-Object {$_.SchemaClassName -eq 'group'})){
        $groupName=[string]$group.Name
        foreach($member in @($group.psbase.Invoke('Members'))){
            try{
                $memberType=$member.GetType()
                $memberName=$memberType.InvokeMember('Name','GetProperty',$null,$member,$null)
                $memberClass=$memberType.InvokeMember('Class','GetProperty',$null,$member,$null)
                $adsPath=$memberType.InvokeMember('AdsPath','GetProperty',$null,$member,$null)
                [pscustomobject][ordered]@{
                    ComputerName=$ComputerName
                    Group=$groupName
                    Name=$(if($adsPath){$adsPath -replace '^WinNT://','' -replace '/','\\'}else{$memberName})
                    ObjectClass=$memberClass
                    PrincipalSource='WinNT'
                }
            }catch{
                Add-DiscoveryDiagnostic 'Local Group Membership' Warning ('{0}: {1}' -f $groupName,$_.Exception.Message)
            }
        }
    }
}

$serviceAccounts=@(
    $services | Where-Object {
        $_.StartName -and
        $_.StartName -notmatch '^(LocalSystem|LocalService|NetworkService|NT AUTHORITY\\|NT SERVICE\\)'
    } | ForEach-Object {
        [pscustomobject][ordered]@{
            ComputerName=$ComputerName
            Account=$_.StartName
            Usage=('Windows service: {0}' -f $_.DisplayName)
            InteractiveLoginRequired='Unknown'
            ManagedServiceAccount=$(if($_.StartName.EndsWith('$')){'Yes'}else{'No'})
        }
    } | Sort-Object Account,Usage -Unique
)
Add-DiscoveryDiagnostic 'Service Accounts' Success 'Collector completed.' $serviceAccounts.Count

$activeDirectory=@()
$adUsers=@()
$adGroupMembership=@()
$adReplication=@()
$dhcpScopes=@()
$dhcpExclusions=@()
$dhcpReservations=@()
$dhcpOptions=@()
$dhcpFailover=@()
$dnsZones=@()

if($basicOnly){
    $reason=('{0} detected; role-specific AD, DHCP, and DNS collectors were skipped.' -f $serverClass)
    foreach($collectorName in 'Active Directory','AD Users','AD Group Membership','AD Replication','DHCP Scopes','DHCP Exclusions','DHCP Reservations','DHCP Options','DHCP Failover','DNS Zones'){
        Add-DiscoveryDiagnostic $collectorName Skipped $reason
    }
}else{
    if($rolesAndFeatures.Name -contains 'AD-Domain-Services'){
        $activeDirectory=Invoke-DiscoveryCollector 'Active Directory' {
            if(-not(Get-Module -ListAvailable ActiveDirectory)){
                throw 'ActiveDirectory module is unavailable.'
            }
            Import-Module ActiveDirectory -ErrorAction Stop
            $dc=Get-ADDomainController -Identity $ComputerName
            $domain=Get-ADDomain
            $forest=Get-ADForest
            [pscustomobject][ordered]@{
                ComputerName=$ComputerName
                FQDN=$dc.HostName
                Site=$dc.Site
                IPv4=$dc.IPv4Address
                OperatingSystem=$dc.OperatingSystem
                GlobalCatalog=$dc.IsGlobalCatalog
                ReadOnly=$dc.IsReadOnly
                DomainFQDN=$domain.DNSRoot
                DomainShortName=$domain.NetBIOSName
                UPNSuffixes=(Join-DiscoveryValue @($forest.UPNSuffixes))
                DomainMode=[string]$domain.DomainMode
                ForestMode=[string]$forest.ForestMode
                PDC=$domain.PDCEmulator
                RIDMaster=$domain.RIDMaster
                InfrastructureMaster=$domain.InfrastructureMaster
                SchemaMaster=$forest.SchemaMaster
                NamingMaster=$forest.DomainNamingMaster
            }
        }

        $adUsers=Invoke-DiscoveryCollector 'AD Users' {
            Import-Module ActiveDirectory -ErrorAction Stop
            $domain=(Get-ADDomain).DNSRoot
            $userProperties='GivenName','Surname','DisplayName','UserPrincipalName','StreetAddress','City','State','PostalCode','co','Title','Department','Company','Manager','Description','Office','OfficePhone','EmailAddress','MobilePhone','Info','Enabled','LastLogonDate','ServicePrincipalName','PasswordNeverExpires','DistinguishedName'
            Get-ADUser -Filter * -Properties $userProperties |
              Sort-Object SamAccountName | ForEach-Object {
                [pscustomobject][ordered]@{
                    Domain=$domain
                    CollectedFrom=$ComputerName
                    DistinguishedName=$_.DistinguishedName
                    FirstName=$_.GivenName
                    LastName=$_.Surname
                    DisplayName=$_.DisplayName
                    SamAccountName=$_.SamAccountName
                    UserPrincipalName=$_.UserPrincipalName
                    Street=$_.StreetAddress
                    City=$_.City
                    State=$_.State
                    PostalCode=$_.PostalCode
                    Country=$_.co
                    JobTitle=$_.Title
                    Department=$_.Department
                    Company=$_.Company
                    Manager=$_.Manager
                    Description=$_.Description
                    Office=$_.Office
                    Telephone=$_.OfficePhone
                    Email=$_.EmailAddress
                    Mobile=$_.MobilePhone
                    Notes=$_.Info
                    Enabled=$_.Enabled
                    LastLogonDate=$_.LastLogonDate
                    LikelyServiceAccount=$(
                        $_.SamAccountName.EndsWith('$') -or
                        @($_.ServicePrincipalName).Count -gt 0
                    )
                    PasswordNeverExpires=$_.PasswordNeverExpires
                }
              }
        }

        $adGroupMembership=Invoke-DiscoveryCollector 'AD Group Membership' {
            Import-Module ActiveDirectory -ErrorAction Stop
            $domain=Get-ADDomain
            $primaryGroups=@{}
            foreach($user in Get-ADUser -Filter * -Properties DisplayName,MemberOf,PrimaryGroupID){
                foreach($groupDn in @($user.MemberOf)){
                    $groupName=(($groupDn -split '(?<!\\),')[0] -replace '^CN=','') -replace '\\,',','
                    [pscustomobject][ordered]@{
                        Domain=$domain.DNSRoot
                        CollectedFrom=$ComputerName
                        Username=$user.SamAccountName
                        Name=$user.DisplayName
                        GroupName=$groupName
                        MembershipType='Direct'
                    }
                }
                $rid=[string]$user.PrimaryGroupID
                if($rid){
                    if(-not $primaryGroups.ContainsKey($rid)){
                        $primarySid='{0}-{1}' -f $domain.DomainSID.Value,$rid
                        $primaryGroups[$rid]=(Get-ADGroup -Identity $primarySid).Name
                    }
                    [pscustomobject][ordered]@{
                        Domain=$domain.DNSRoot
                        CollectedFrom=$ComputerName
                        Username=$user.SamAccountName
                        Name=$user.DisplayName
                        GroupName=$primaryGroups[$rid]
                        MembershipType='Primary'
                    }
                }
            }
        }

        $adReplication=Invoke-DiscoveryCollector 'AD Replication' {
            Import-Module ActiveDirectory -ErrorAction Stop
            Get-ADReplicationPartnerMetadata -Target $ComputerName -Scope Server |
              ForEach-Object {
                [pscustomobject][ordered]@{
                    ComputerName=$ComputerName
                    FromServer=$_.Partner
                    ToServer=$_.Server
                    LastSync=$_.LastReplicationSuccess
                    Status=$(if($_.LastReplicationResult -eq 0){
                        'Success'
                    }else{
                        'Error {0}' -f $_.LastReplicationResult
                    })
                }
              }
        }
    }else{
        foreach($collectorName in 'Active Directory','AD Users','AD Group Membership','AD Replication'){
            Add-DiscoveryDiagnostic $collectorName Skipped 'AD DS role is not installed.'
        }
    }

    if($rolesAndFeatures.Name -contains 'DHCP'){
        $dhcpScopes=Invoke-DiscoveryCollector 'DHCP Scopes' {
            if(-not(Get-Command Get-DhcpServerv4Scope -ErrorAction SilentlyContinue)){
                throw 'DHCP cmdlets are unavailable.'
            }
            foreach($scope in Get-DhcpServerv4Scope -ComputerName $ComputerName -ErrorAction Stop){
                $stats=try{
                    Get-DhcpServerv4ScopeStatistics -ComputerName $ComputerName -ScopeId $scope.ScopeId -ErrorAction Stop
                }catch{
                    Add-DiscoveryDiagnostic 'DHCP Scope Statistics' Warning ('{0}: {1}' -f $scope.ScopeId,$_.Exception.Message)
                    $null
                }
                [pscustomobject][ordered]@{
                    ServerName=$ComputerName
                    ScopeName=$scope.Name
                    Status=$scope.State
                    ScopeId=$scope.ScopeId
                    StartRange=$scope.StartRange
                    EndRange=$scope.EndRange
                    SubnetMask=$scope.SubnetMask
                    AddressPool=('{0} - {1}' -f $scope.StartRange,$scope.EndRange)
                    LeaseDuration=$scope.LeaseDuration
                    Utilization=$(if($stats){$stats.PercentageInUse}else{$null})
                    NAP=$(if($scope.PSObject.Properties['NapEnable']){$scope.NapEnable}else{$null})
                }
            }
        }

        $dhcpExclusions=Invoke-DiscoveryCollector 'DHCP Exclusions' {
            foreach($scope in $dhcpScopes){
                Get-DhcpServerv4ExclusionRange -ComputerName $ComputerName -ScopeId $scope.ScopeId -ErrorAction Stop | ForEach-Object {
                    [pscustomobject][ordered]@{
                        ServerName=$ComputerName
                        ScopeId=$scope.ScopeId
                        ScopeName=$scope.ScopeName
                        StartRange=$_.StartRange
                        EndRange=$_.EndRange
                    }
                  }
            }
        }

        $dhcpReservations=Invoke-DiscoveryCollector 'DHCP Reservations' {
            foreach($scope in $dhcpScopes){
                Get-DhcpServerv4Reservation -ComputerName $ComputerName -ScopeId $scope.ScopeId -ErrorAction Stop | ForEach-Object {
                    [pscustomobject][ordered]@{
                        ServerName=$ComputerName
                        ScopeId=$scope.ScopeId
                        ScopeName=$scope.ScopeName
                        IPAddress=$_.IPAddress
                        ClientId=$_.ClientId
                        Name=$_.Name
                        Description=$_.Description
                        Type=$_.Type
                    }
                  }
            }
        }

        $dhcpOptions=Invoke-DiscoveryCollector 'DHCP Options' {
            foreach($scope in $dhcpScopes){
                Get-DhcpServerv4OptionValue -ComputerName $ComputerName -ScopeId $scope.ScopeId -ErrorAction Stop | ForEach-Object {
                    [pscustomobject][ordered]@{
                        ServerName=$ComputerName
                        ScopeId=$scope.ScopeId
                        ScopeName=$scope.ScopeName
                        OptionId=$_.OptionId
                        Name=$_.Name
                        Type=$_.Type
                        Value=(Join-DiscoveryValue @($_.Value) ', ')
                        VendorClass=$_.VendorClass
                        UserClass=$_.UserClass
                        PolicyName=$_.PolicyName
                    }
                  }
            }
        }

        $dhcpFailover=Invoke-DiscoveryCollector 'DHCP Failover' {
            Get-DhcpServerv4Failover -ComputerName $ComputerName -ErrorAction SilentlyContinue | ForEach-Object {
                [pscustomobject][ordered]@{
                    ServerName=$ComputerName
                    Name=$_.Name
                    PartnerServer=$_.PartnerServer
                    Mode=$_.Mode
                    State=$_.State
                    ServerRole=$_.ServerRole
                    ScopeId=(Join-DiscoveryValue @($_.ScopeId))
                    LoadBalancePercent=$_.LoadBalancePercent
                    ReservePercent=$_.ReservePercent
                    MaxClientLeadTime=$_.MaxClientLeadTime
                    AutoStateTransition=$_.AutoStateTransition
                    StateSwitchInterval=$_.StateSwitchInterval
                }
              }
        }
    }else{
        foreach($collectorName in 'DHCP Scopes','DHCP Exclusions','DHCP Reservations','DHCP Options','DHCP Failover'){
            Add-DiscoveryDiagnostic $collectorName Skipped 'DHCP role is not installed.'
        }
    }

    if($rolesAndFeatures.Name -contains 'DNS'){
        $dnsZones=Invoke-DiscoveryCollector 'DNS Zones' {
            if(-not(Get-Command Get-DnsServerZone -ErrorAction SilentlyContinue)){
                throw 'DNS cmdlets are unavailable.'
            }
            Get-DnsServerZone -ComputerName $ComputerName -ErrorAction Stop |
              Sort-Object ZoneName | Select-Object @{n='ComputerName';e={$ComputerName}},ZoneName,ZoneType,IsReverseLookupZone,IsDsIntegrated,DynamicUpdate,ReplicationScope
        }
    }else{
        Add-DiscoveryDiagnostic 'DNS Zones' Skipped 'DNS role is not installed.'
    }
}

[pscustomobject][ordered]@{
    Metadata=[pscustomobject][ordered]@{
        ComputerName=$ComputerName
        ServerClass=$serverClass
        CollectionMode=$collectionMode
        IsSqlServer=$isSqlServer
        IsExchangeServer=$isExchangeServer
        HasActiveDirectory=($rolesAndFeatures.Name -contains 'AD-Domain-Services')
        HasDhcp=($rolesAndFeatures.Name -contains 'DHCP')
        HasDns=($rolesAndFeatures.Name -contains 'DNS')
        CollectedAt=(Get-Date)
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
    ActiveDirectory=@($activeDirectory)
    ADUsers=@($adUsers)
    ADGroupMembership=@($adGroupMembership)
    ADReplication=@($adReplication)
    DHCPScopes=@($dhcpScopes)
    DHCPExclusions=@($dhcpExclusions)
    DHCPReservations=@($dhcpReservations)
    DHCPOptions=@($dhcpOptions)
    DHCPFailover=@($dhcpFailover)
    DNSZones=@($dnsZones)
    # New-Object keeps this collector compatible with PowerShell 4.0. Use
    # ToArray(): @($list) can throw on its wrapped List[object] in PowerShell 5+.
    Diagnostics=$script:Diagnostics.ToArray()
}
