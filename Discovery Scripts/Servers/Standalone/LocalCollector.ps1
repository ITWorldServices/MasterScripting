function Get-LocalBasicDiscoveryData {
    param([bool]$IncludeMicrosoftTasks=$false)
    $ErrorActionPreference='Stop'
    $ComputerName=$env:COMPUTERNAME
    $script:LocalDiagnostics=New-Object 'System.Collections.Generic.List[object]'
    # No -ComputerName or -Credential: all WMI calls execute on this server.
    $system=Invoke-LocalDiscoverySection 'System' {
        $cs=Get-WmiObject Win32_ComputerSystem
        $os=Get-WmiObject Win32_OperatingSystem
        $cpu=@(Get-WmiObject Win32_Processor)
        $bios=Get-WmiObject Win32_BIOS
        $virtual=$cs.Manufacturer -match 'VMware|QEMU|Xen|innotek' -or $cs.Model -match 'Virtual Machine|VirtualBox|KVM|HVM domU|VMware|Bochs|Xen'
        New-LocalDiscoveryRow @{
            ComputerName=$ComputerName; Domain=$cs.Domain; Manufacturer=$cs.Manufacturer
            Model=$cs.Model; SerialNumber=$bios.SerialNumber
            PhysicalOrVirtual=$(if($virtual){'Virtual'}else{'Physical'})
            OperatingSystem=$os.Caption; Version=$os.Version; Build=$os.BuildNumber
            LastBoot=(Convert-LocalDiscoveryDate $os.LastBootUpTime)
            CPU=(Join-LocalDiscoveryValue @($cpu | ForEach-Object {$_.Name}))
            Sockets=$cpu.Count; Cores=($cpu | Measure-Object NumberOfCores -Sum).Sum
            LogicalProcessors=($cpu | Measure-Object NumberOfLogicalProcessors -Sum).Sum
            MemoryGB=[math]::Round($cs.TotalPhysicalMemory/1GB,2)
        }
    }
    $network=Invoke-LocalDiscoverySection 'Network' {
        foreach($adapter in @(Get-WmiObject Win32_NetworkAdapterConfiguration -Filter 'IPEnabled=True')){
            $addresses=@($adapter.IPAddress)
            for($i=0;$i -lt $addresses.Count;$i++){
                if([string]$addresses[$i] -notmatch '^\d{1,3}(\.\d{1,3}){3}$'){continue}
                $mask=$null; $prefix=$null
                if(@($adapter.IPSubnet).Count -gt $i){$mask=@($adapter.IPSubnet)[$i]}
                if($mask -match '^\d{1,3}(\.\d{1,3}){3}$'){
                    $bits=($mask.Split('.') | ForEach-Object {[Convert]::ToString([int]$_,2).PadLeft(8,'0')}) -join ''
                    if($bits -match '^1*0*$'){$prefix=($bits -replace '0','').Length}
                }
                New-LocalDiscoveryRow @{
                    ComputerName=$ComputerName; Interface=$adapter.Description; IPAddress=$addresses[$i]
                    PrefixLength=$prefix; SubnetMask=$mask
                    Gateway=(Join-LocalDiscoveryValue @($adapter.DefaultIPGateway))
                    DNS=(Join-LocalDiscoveryValue @($adapter.DNSServerSearchOrder)); MacAddress=$adapter.MACAddress
                }
            }
        }
    }
    $storage=Invoke-LocalDiscoverySection 'Storage' {
        Get-WmiObject Win32_LogicalDisk -Filter 'DriveType=3' | Sort-Object DeviceID | ForEach-Object {
            New-LocalDiscoveryRow @{
                ComputerName=$ComputerName; Drive=$_.DeviceID; Label=$_.VolumeName; FileSystem=$_.FileSystem
                SizeGB=[math]::Round($_.Size/1GB,2); FreeGB=[math]::Round($_.FreeSpace/1GB,2)
                PercentFree=$(if($_.Size){[math]::Round(100*$_.FreeSpace/$_.Size,2)}else{$null})
            }
        }
    }
    $roles=Invoke-LocalDiscoverySection 'Roles and Features' {
        Get-WmiObject Win32_ServerFeature | Sort-Object Name | ForEach-Object {
            New-LocalDiscoveryRow @{ComputerName=$ComputerName; Name=$_.Name; DisplayName=$_.Name; FeatureType='Legacy server feature'}
        }
    }
    $applications=Invoke-LocalDiscoverySection 'Applications' {
        # Local registry reads avoid the slow per-property remote registry queries.
        foreach($path in 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall','HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall'){
            if(-not(Test-Path $path)){continue}
            foreach($key in Get-ChildItem $path){
                try{
                    $app=Get-ItemProperty $key.PSPath
                    if(-not $app.DisplayName -or $app.SystemComponent -eq 1){continue}
                    New-LocalDiscoveryRow @{
                        ComputerName=$ComputerName; Name=$app.DisplayName; Version=$app.DisplayVersion
                        Publisher=$app.Publisher; InstallDate=$app.InstallDate; InstallLocation=$app.InstallLocation
                    }
                }catch{Add-LocalDiscoveryDiagnostic 'Applications' Warning ('{0}: {1}' -f $key.PSChildName,$_.Exception.Message)}
            }
        }
    }
    $services=Invoke-LocalDiscoverySection 'Services' {
        Get-WmiObject Win32_Service | Sort-Object DisplayName | ForEach-Object {
            New-LocalDiscoveryRow @{
                ComputerName=$ComputerName; Name=$_.Name; DisplayName=$_.DisplayName; State=$_.State
                StartMode=$_.StartMode; StartName=$_.StartName; PathName=$_.PathName
            }
        }
    }
    $isSql=@($services | Where-Object {$_.Name -match '^(MSSQLSERVER|MSSQL\$|SQLSERVERAGENT|SQLAgent\$|MSOLAP\$|MSSQLServerOLAPService|ReportServer)'}).Count -gt 0
    $isExchange=@($services | Where-Object {$_.Name -match '^MSExchange'}).Count -gt 0
    $class=if($isSql -and $isExchange){'SQL and Exchange'}elseif($isSql){'SQL Server'}elseif($isExchange){'Exchange Server'}else{'General Server'}
    $hasAd=@($services | Where-Object {$_.Name -eq 'NTDS'}).Count -gt 0
    $hasDns=@($services | Where-Object {$_.Name -eq 'DNS'}).Count -gt 0
    $hasDhcp=@($services | Where-Object {$_.Name -eq 'DHCPServer'}).Count -gt 0
    # Local domain-role check prevents domain users being presented as local accounts.
    $isDc=$hasAd
    try{$isDc=[int](Get-WmiObject Win32_ComputerSystem).DomainRole -ge 4}
    catch{Add-LocalDiscoveryDiagnostic 'Domain Role' Warning 'Unable to determine domain-controller status; local account collection will be skipped.'; $isDc=$true}
    foreach($row in $system){
        $row | Add-Member NoteProperty ServerClass $class -Force
        $row | Add-Member NoteProperty CollectionMode 'Local legacy basic only' -Force
    }
    $ports=Invoke-LocalDiscoverySection 'Listening Ports' {
        # TCP listeners have an unspecified foreign endpoint; no localized state-name test.
        $lines=@(& "$env:SystemRoot\System32\netstat.exe" -ano -p tcp)
        if($LASTEXITCODE -ne 0){throw ('netstat exited with code {0}' -f $LASTEXITCODE)}
        foreach($line in $lines){
            $fields=@(($line.Trim() -split '\s+'))
            if($fields.Count -lt 5 -or $fields[0] -ne 'TCP' -or $fields[2] -notmatch '^(0\.0\.0\.0|\[::\]):0$'){continue}
            if($fields[1] -notmatch '^(.*):(\d+)$'){continue}
            $address=$matches[1].Trim('[',']'); $port=[int]$matches[2]; $processId=[int]$fields[-1]
            $processName=$null
            try{$processName=(Get-Process -Id $processId -ErrorAction Stop).ProcessName}catch{}
            New-LocalDiscoveryRow @{ComputerName=$ComputerName; LocalAddress=$address; LocalPort=$port; ProcessId=$processId; ProcessName=$processName}
        }
    }
    $shares=Invoke-LocalDiscoverySection 'Shares and Permissions' {
        foreach($share in Get-WmiObject Win32_Share){
            $sharePermissions=@(); $ntfsPermissions=@()
            try{
                $security=Get-WmiObject Win32_LogicalShareSecuritySetting | Where-Object {$_.Name -eq $share.Name} | Select-Object -First 1
                if(-not $security){throw 'Share security descriptor is unavailable.'}
                $descriptor=$security.GetSecurityDescriptor()
                if($descriptor.ReturnValue -ne 0){throw ('GetSecurityDescriptor returned {0}' -f $descriptor.ReturnValue)}
                foreach($ace in @($descriptor.Descriptor.DACL)){
                    $identity=$ace.Trustee.SIDString
                    if($ace.Trustee.Name){$identity=('{0}\{1}' -f $ace.Trustee.Domain,$ace.Trustee.Name).TrimStart('\')}
                    $rights=[string][Security.AccessControl.FileSystemRights]$ace.AccessMask
                    $kind=if($ace.AceType -eq 0){'Allow'}elseif($ace.AceType -eq 1){'Deny'}else{'ACE type '+$ace.AceType}
                    $sharePermissions+=@('{0}: {1} ({2})' -f $identity,$rights,$kind)
                }
                if($null -eq $descriptor.Descriptor.DACL){$sharePermissions=@('Null DACL (unrestricted)')}
            }catch{Add-LocalDiscoveryDiagnostic 'Shares and Permissions' Warning ('Share {0}: {1}' -f $share.Name,$_.Exception.Message)}
            try{
                if($share.Path -and (Test-Path -LiteralPath $share.Path)){
                    $ntfsPermissions=@((Get-Acl -LiteralPath $share.Path).Access | ForEach-Object {'{0}: {1} ({2})' -f $_.IdentityReference,$_.FileSystemRights,$_.AccessControlType})
                }
            }catch{Add-LocalDiscoveryDiagnostic 'Shares and Permissions' Warning ('NTFS {0}: {1}' -f $share.Name,$_.Exception.Message)}
            New-LocalDiscoveryRow @{
                ComputerName=$ComputerName; Name=$share.Name; Path=$share.Path; UNCPath=('\\{0}\{1}' -f $ComputerName,$share.Name)
                Description=$share.Description; Special=([int64]$share.Type -ge 2147483648)
                SharePermissions=(Join-LocalDiscoveryValue $sharePermissions); NTFSPermissions=(Join-LocalDiscoveryValue $ntfsPermissions)
            }
        }
    }
    $tasks=Invoke-LocalDiscoverySection 'Scheduled Tasks' {
        $scheduler=New-Object -ComObject 'Schedule.Service'
        $scheduler.Connect()
        $folders=New-Object 'System.Collections.Generic.Queue[object]'
        $folders.Enqueue($scheduler.GetFolder('\'))
        while($folders.Count -gt 0){
            $folder=$folders.Dequeue()
            try{
                foreach($child in $folder.GetFolders(0)){
                    if(-not $IncludeMicrosoftTasks -and ($child.Path -eq '\Microsoft' -or $child.Path -like '\Microsoft\*')){continue}
                    $folders.Enqueue($child)
                }
                foreach($task in $folder.GetTasks(1)){
                    $definition=$task.Definition
                    $actions=@($definition.Actions | ForEach-Object {
                        if($_.Type -eq 0){('{0} {1}' -f $_.Path,$_.Arguments).Trim()}else{'Task action type '+$_.Type}
                    })
                    $stateNames=@('Unknown','Disabled','Queued','Ready','Running')
                    $state=[string]$task.State
                    if([int]$task.State -ge 0 -and [int]$task.State -lt $stateNames.Count){$state=$stateNames[[int]$task.State]}
                    New-LocalDiscoveryRow @{
                        ComputerName=$ComputerName; Name=$task.Name; TaskPath=$folder.Path; Executable=(Join-LocalDiscoveryValue $actions)
                        State=$state; RunAsUser=$definition.Principal.UserId; Description=$definition.RegistrationInfo.Description
                    }
                }
            }catch{Add-LocalDiscoveryDiagnostic 'Scheduled Tasks' Warning ('{0}: {1}' -f $folder.Path,$_.Exception.Message)}
        }
    }
    $accounts=@(); $membership=@()
    if($isDc){
        Add-LocalDiscoveryDiagnostic 'Local Accounts' Skipped 'Local accounts are not collected on a domain controller or when its role is unknown.'
        Add-LocalDiscoveryDiagnostic 'Local Group Membership' Skipped 'Local groups are not collected on a domain controller or when its role is unknown.'
    }else{
        $accounts=Invoke-LocalDiscoverySection 'Local Accounts' {
            Get-WmiObject Win32_UserAccount -Filter 'LocalAccount=True' | ForEach-Object {
                New-LocalDiscoveryRow @{ComputerName=$ComputerName; Name=$_.Name; Enabled=(-not $_.Disabled); Description=$_.Description; LastLogon=$null; PasswordExpires=$_.PasswordExpires}
            }
            Add-LocalDiscoveryDiagnostic 'Local Accounts' Warning 'Win32_UserAccount does not expose LastLogon; that field is blank.'
        }
        $membership=Invoke-LocalDiscoverySection 'Local Group Membership' {
            $computer=[ADSI]('WinNT://{0},computer' -f $ComputerName)
            foreach($group in @($computer.psbase.Children | Where-Object {$_.SchemaClassName -eq 'group'})){
                $groupName=[string]$group.Name
                try{
                    foreach($member in @($group.psbase.Invoke('Members'))){
                        $memberType=$member.GetType()
                        $path=$memberType.InvokeMember('AdsPath','GetProperty',$null,$member,$null)
                        $kind=$memberType.InvokeMember('Class','GetProperty',$null,$member,$null)
                        New-LocalDiscoveryRow @{ComputerName=$ComputerName; Group=$groupName; Name=($path -replace '^WinNT://','' -replace '/','\'); ObjectClass=$kind; PrincipalSource='WinNT'}
                    }
                }catch{Add-LocalDiscoveryDiagnostic 'Local Group Membership' Warning ('{0}: {1}' -f $groupName,$_.Exception.Message)}
            }
        }
    }
    $serviceAccounts=Invoke-LocalDiscoverySection 'Service Accounts' {
        $services | Where-Object {$_.StartName -and $_.StartName -notmatch '^(LocalSystem|LocalService|NetworkService|NT AUTHORITY\\|NT SERVICE\\)'} | ForEach-Object {
            New-LocalDiscoveryRow @{ComputerName=$ComputerName; Account=$_.StartName; Usage=('Windows service: {0}' -f $_.DisplayName); InteractiveLoginRequired='Unknown'; ManagedServiceAccount=$(if($_.StartName.EndsWith('$')){'Yes'}else{'No'})}
        }
    }
    $payload=New-LocalDiscoveryRow @{
        Metadata=(New-LocalDiscoveryRow @{ComputerName=$ComputerName; ServerClass=$class; CollectionMode='Local legacy basic only'; IsSqlServer=$isSql; IsExchangeServer=$isExchange; HasActiveDirectory=$hasAd; HasDns=$hasDns; HasDhcp=$hasDhcp; CollectedAt=Get-Date})
        System=@($system); Network=@($network); Storage=@($storage); RolesAndFeatures=@($roles)
        Applications=@($applications); Services=@($services); ListeningPorts=@($ports)
        SharesAndPermissions=@($shares); ScheduledTasks=@($tasks); LocalAccounts=@($accounts)
        LocalGroupMembership=@($membership); ServiceAccounts=@($serviceAccounts)
    }
    foreach($name in 'ActiveDirectory','ADUsers','ADGroupMembership','ADReplication','DHCPScopes','DHCPExclusions','DHCPReservations','DHCPOptions','DHCPFailover','DNSZones'){
        $payload | Add-Member NoteProperty $name @()
        $reason='Legacy local collection supports basic inventory only; role-specific details require the modern collector.'
        if($isSql -or $isExchange){$reason=$class+' detected; role-specific collectors are excluded.'}
        elseif(($name -like 'AD*' -or $name -eq 'ActiveDirectory') -and -not $hasAd){$reason='AD DS is not installed.'}
        elseif($name -like 'DHCP*' -and -not $hasDhcp){$reason='DHCP server is not installed.'}
        elseif($name -eq 'DNSZones' -and -not $hasDns){$reason='DNS server is not installed.'}
        Add-LocalDiscoveryDiagnostic $name Skipped $reason
    }
    $payload | Add-Member NoteProperty Diagnostics $script:LocalDiagnostics.ToArray()
    return $payload
}
