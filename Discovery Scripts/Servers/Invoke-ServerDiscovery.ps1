#Requires -Version 5.1
<#
.SYNOPSIS
Collects discovery data from the local Windows Server and writes a timestamped
copy of the TPT Server Migration Planning workbook.

.DESCRIPTION
Phase 1 is local-only. It does not enumerate AD computers, use remoting, install
modules, call Win32_Product, or recursively scan file shares. Each collector is
isolated; failures are recorded on the Diagnostics worksheet.
#>
[CmdletBinding()]
param(
    [string]$TemplatePath,
    [string]$OutputDirectory = 'C:\Temp',
    [switch]$IncludeMicrosoftTasks
)

$ErrorActionPreference = 'Stop'
$ProgressPreference = 'SilentlyContinue'
$script:Diagnostics = [System.Collections.Generic.List[object]]::new()
$ComputerName = $env:COMPUTERNAME

function Add-Diagnostic {
    param([string]$Collector,[string]$Status,[string]$Message,[int]$Count = 0)
    $script:Diagnostics.Add([pscustomobject][ordered]@{
        Time=Get-Date; Collector=$Collector; Status=$Status; RecordCount=$Count; Message=$Message
    })
}

function Invoke-Collector {
    param([string]$Name,[scriptblock]$Action)
    try {
        $result = @(& $Action)
        Add-Diagnostic $Name 'Success' 'Collector completed.' $result.Count
        Write-Output -NoEnumerate $result
        return
    } catch {
        Add-Diagnostic $Name 'Failed' $_.Exception.Message
        Write-Warning ('{0}: {1}' -f $Name,$_.Exception.Message)
        Write-Output -NoEnumerate @()
        return
    }
}

function Join-Unique {
    param([object[]]$Values,[string]$Separator = '; ')
    return (($Values | Where-Object { $null -ne $_ -and "$_".Trim() } |
        ForEach-Object { "$_".Trim() } | Sort-Object -Unique) -join $Separator)
}

function Convert-PrefixToMask {
    param([ValidateRange(0,32)][int]$Prefix)
    $bits = (('1' * $Prefix) -join '').PadRight(32,'0')
    return ((0,8,16,24 | ForEach-Object {
        [Convert]::ToInt32($bits.Substring($_,8),2)
    }) -join '.')
}

function Set-TemplateRow {
    param($Worksheet,[int]$HeaderRow,[int]$DataRow,[hashtable]$Values)
    if (-not $Worksheet -or -not $Worksheet.Dimension) { return }
    for ($column=1; $column -le $Worksheet.Dimension.End.Column; $column++) {
        $header = ([string]$Worksheet.Cells[$HeaderRow,$column].Text).Trim()
        if ($header -and $Values.ContainsKey($header)) {
            $Worksheet.Cells[$DataRow,$column].Value = $Values[$header]
        }
    }
}

function New-YesNoTemplateRow {
    param($Worksheet,[int]$HeaderRow)
    $values=@{'Server Name'=$ComputerName}
    if(-not $Worksheet -or -not $Worksheet.Dimension){ return $values }
    for($column=1;$column -le $Worksheet.Dimension.End.Column;$column++){
        $header=([string]$Worksheet.Cells[$HeaderRow,$column].Text).Trim()
        if($header -and $header -ne 'Server Name'){ $values[$header]='No' }
    }
    return $values
}

function Clear-TemplateValues {
    param($Worksheet,[string]$Address)
    if(-not $Worksheet){ return }
    foreach($cell in $Worksheet.Cells[$Address]){ $cell.Value=$null }
}

function Add-DataSheet {
    param($Workbook,[string]$Name,[object[]]$Rows)
    if ($Workbook.Worksheets[$Name]) { $Workbook.Worksheets.Delete($Name) }
    $sheet = $Workbook.Worksheets.Add($Name)
    $data = @($Rows)
    if (-not $data.Count) {
        $sheet.Cells[1,1].Value = 'Status'
        $sheet.Cells[2,1].Value = 'No records returned.'
        return
    }
    $headers = @($data[0].PSObject.Properties.Name)
    for ($column=0; $column -lt $headers.Count; $column++) {
        $sheet.Cells[1,($column+1)].Value = $headers[$column]
        $sheet.Cells[1,($column+1)].Style.Font.Bold = $true
    }
    for ($row=0; $row -lt $data.Count; $row++) {
        for ($column=0; $column -lt $headers.Count; $column++) {
            $value = $data[$row].PSObject.Properties[$headers[$column]].Value
            if ($value -is [array]) { $value = Join-Unique $value }
            $cell = $sheet.Cells[($row+2),($column+1)]
            if ($value -is [bool]) { $value = if ($value) { 'Yes' } else { 'No' } }
            $cell.Value = $value
            if ($value -is [datetime]) { $cell.Style.Numberformat.Format = 'yyyy-mm-dd HH:mm:ss' }
        }
    }
    $sheet.View.FreezePanes(2,1)
    $sheet.Cells[$sheet.Dimension.Address].AutoFitColumns()
}

if ($env:OS -ne 'Windows_NT') { throw 'Run this script on the Windows Server being inventoried.' }
function Initialize-ImportExcel {
    $requiredVersion = '7.8.10'

    if (Get-Module -ListAvailable -Name ImportExcel |
        Where-Object { $_.Version -ge [version]$requiredVersion } |
        Select-Object -First 1) {
        Import-Module ImportExcel -MinimumVersion $requiredVersion -ErrorAction Stop
        Add-Diagnostic 'ImportExcel Bootstrap' Success ('ImportExcel {0} or later is already installed.' -f $requiredVersion)
        return
    }

    Write-Host ('ImportExcel {0} is not installed. Bootstrapping PowerShell Gallery access...' -f $requiredVersion) -ForegroundColor Yellow

    try {
        [Net.ServicePointManager]::SecurityProtocol =
            [Net.ServicePointManager]::SecurityProtocol -bor [Net.SecurityProtocolType]::Tls12
        Add-Diagnostic 'ImportExcel Bootstrap' Success 'TLS 1.2 enabled for this PowerShell process.'

        $nuget = Get-PackageProvider -Name NuGet -ListAvailable -ErrorAction SilentlyContinue |
            Where-Object { $_.Version -ge [version]'2.8.5.201' } |
            Select-Object -First 1

        if (-not $nuget) {
            Write-Host 'Installing the NuGet package provider...' -ForegroundColor Yellow
            Install-PackageProvider -Name NuGet -MinimumVersion 2.8.5.201 -Force -Confirm:$false | Out-Null
            Add-Diagnostic 'ImportExcel Bootstrap' Success 'NuGet package provider installed.'
        }

        $gallery = Get-PSRepository -Name PSGallery -ErrorAction SilentlyContinue
        if (-not $gallery) {
            Write-Host 'Registering the default PowerShell Gallery repository...' -ForegroundColor Yellow
            Register-PSRepository -Default -ErrorAction Stop
            $gallery = Get-PSRepository -Name PSGallery -ErrorAction Stop
            Add-Diagnostic 'ImportExcel Bootstrap' Success 'PowerShell Gallery repository registered.'
        }

        $originalPolicy = $gallery.InstallationPolicy
        if ($originalPolicy -ne 'Trusted') {
            Set-PSRepository -Name PSGallery -InstallationPolicy Trusted -ErrorAction Stop
        }

        try {
            Write-Host ('Installing ImportExcel {0} for the current user...' -f $requiredVersion) -ForegroundColor Yellow
            Install-Module -Name ImportExcel -RequiredVersion $requiredVersion -Repository PSGallery -Scope CurrentUser -Force -AllowClobber -Confirm:$false -ErrorAction Stop
        }
        finally {
            if ($originalPolicy -and $originalPolicy -ne 'Trusted') {
                Set-PSRepository -Name PSGallery -InstallationPolicy $originalPolicy -ErrorAction SilentlyContinue
            }
        }

        Import-Module ImportExcel -RequiredVersion $requiredVersion -Force -ErrorAction Stop
        Add-Diagnostic 'ImportExcel Bootstrap' Success ('ImportExcel {0} installed and imported.' -f $requiredVersion)
    }
    catch {
        Add-Diagnostic 'ImportExcel Bootstrap' Failed $_.Exception.Message
        throw ('Unable to install ImportExcel automatically. Confirm HTTPS access to www.powershellgallery.com and the NuGet endpoints, plus any required proxy configuration. Error: {0}' -f $_.Exception.Message)
    }
}

Initialize-ImportExcel
if ([string]::IsNullOrWhiteSpace($TemplatePath)) {
    $templateCandidates = @(
        (Join-Path $PSScriptRoot 'TPT - Server Migration Planning Document.xlsx'),
        (Join-Path $PSScriptRoot '..\..\TPT - Server Migration Planning Document.xlsx')
    )
    $TemplatePath = $templateCandidates | Where-Object { Test-Path -LiteralPath $_ } | Select-Object -First 1
    if (-not $TemplatePath) {
        throw 'Workbook template not found. Place TPT - Server Migration Planning Document.xlsx beside this script or provide -TemplatePath.'
    }
}
$template = Get-Item -LiteralPath $TemplatePath
if ($template.Extension -ne '.xlsx') { throw 'TemplatePath must be an .xlsx file.' }
if (-not (Test-Path -LiteralPath $OutputDirectory)) {
    New-Item -ItemType Directory -Path $OutputDirectory -Force | Out-Null
}
$outputPath = Join-Path $OutputDirectory ('{0}-ServerDiscovery-{1}.xlsx' -f $ComputerName,(Get-Date -Format yyyyMMdd-HHmmss))
Copy-Item -LiteralPath $template.FullName -Destination $outputPath -Force

try {
    $identity = [Security.Principal.WindowsIdentity]::GetCurrent()
    $principal = [Security.Principal.WindowsPrincipal]::new($identity)
    if ($principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
        Add-Diagnostic Prerequisites Success 'Running elevated.'
    } else {
        Add-Diagnostic Prerequisites Warning 'Not elevated; some data may be incomplete.'
    }
} catch { Add-Diagnostic Prerequisites Warning 'Unable to determine elevation.' }

$system = Invoke-Collector 'System' {
    $cs=Get-CimInstance Win32_ComputerSystem
    $os=Get-CimInstance Win32_OperatingSystem
    $cpu=@(Get-CimInstance Win32_Processor)
    $bios=Get-CimInstance Win32_BIOS
    $virtual=$cs.HypervisorPresent -or $cs.Model -match 'Virtual|VMware|KVM|HVM|VirtualBox'
    [pscustomobject][ordered]@{
        ComputerName=$cs.Name; Domain=$cs.Domain; Manufacturer=$cs.Manufacturer; Model=$cs.Model
        SerialNumber=$bios.SerialNumber; PhysicalOrVirtual=$(if($virtual){'Virtual'}else{'Physical'})
        OperatingSystem=$os.Caption; Version=$os.Version; Build=$os.BuildNumber; LastBoot=$os.LastBootUpTime
        CPU=Join-Unique $cpu.Name; Sockets=$cpu.Count
        Cores=($cpu|Measure-Object NumberOfCores -Sum).Sum
        LogicalProcessors=($cpu|Measure-Object NumberOfLogicalProcessors -Sum).Sum
        MemoryGB=[math]::Round($cs.TotalPhysicalMemory/1GB,2)
    }
}

$network = Invoke-Collector 'Network' {
    foreach($config in Get-NetIPConfiguration | Where-Object {$_.NetAdapter.Status -eq 'Up'}) {
        foreach($ip in @($config.IPv4Address)) {
            [pscustomobject][ordered]@{
                ComputerName=$ComputerName; Interface=$config.InterfaceAlias; IPAddress=$ip.IPAddress
                PrefixLength=$ip.PrefixLength; SubnetMask=Convert-PrefixToMask $ip.PrefixLength
                Gateway=Join-Unique $config.IPv4DefaultGateway.NextHop
                DNS=Join-Unique @($config.DNSServer.ServerAddresses | Where-Object {
                    ($_ -as [ipaddress]) -and ([ipaddress]$_).AddressFamily -eq [Net.Sockets.AddressFamily]::InterNetwork
                })
                MacAddress=$config.NetAdapter.MacAddress
            }
        }
    }
}

$disks = Invoke-Collector 'Disks' {
    Get-CimInstance Win32_LogicalDisk -Filter 'DriveType=3' | Sort-Object DeviceID | ForEach-Object {
        [pscustomobject][ordered]@{
            ComputerName=$ComputerName; Drive=$_.DeviceID; Label=$_.VolumeName; FileSystem=$_.FileSystem
            SizeGB=[math]::Round($_.Size/1GB,2); FreeGB=[math]::Round($_.FreeSpace/1GB,2)
            PercentFree=$(if($_.Size){[math]::Round(100*$_.FreeSpace/$_.Size,2)}else{$null})
        }
    }
}

$ports = Invoke-Collector 'Listening Ports' {
    Get-NetTCPConnection -State Listen | Sort-Object LocalPort,LocalAddress -Unique | ForEach-Object {
        [pscustomobject][ordered]@{
            ComputerName=$ComputerName; LocalAddress=$_.LocalAddress; LocalPort=$_.LocalPort
            ProcessId=$_.OwningProcess
            ProcessName=$(try{(Get-Process -Id $_.OwningProcess -ErrorAction Stop).ProcessName}catch{$null})
        }
    }
}

$applications = Invoke-Collector 'Applications' {
    $paths=@(
      'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*',
      'HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*'
    )
    Get-ItemProperty $paths -ErrorAction SilentlyContinue |
      Where-Object {$_.DisplayName -and $_.SystemComponent -ne 1} |
      Sort-Object DisplayName,DisplayVersion -Unique | ForEach-Object {
        [pscustomobject][ordered]@{
            ComputerName=$ComputerName; Name=$_.DisplayName; Version=$_.DisplayVersion
            Publisher=$_.Publisher; InstallDate=$_.InstallDate; InstallLocation=$_.InstallLocation
        }
      }
}

$services = Invoke-Collector 'Services' {
    Get-CimInstance Win32_Service | Sort-Object DisplayName | ForEach-Object {
        [pscustomobject][ordered]@{
            ComputerName=$ComputerName; Name=$_.Name; DisplayName=$_.DisplayName
            State=$_.State; StartMode=$_.StartMode; StartName=$_.StartName; PathName=$_.PathName
        }
    }
}

$features = Invoke-Collector 'Roles and Features' {
    if(-not(Get-Command Get-WindowsFeature -ErrorAction SilentlyContinue)){throw 'Get-WindowsFeature is unavailable.'}
    Get-WindowsFeature | Where-Object Installed | ForEach-Object {
        [pscustomobject][ordered]@{
            ComputerName=$ComputerName; Name=$_.Name; DisplayName=$_.DisplayName; FeatureType=$_.FeatureType
        }
    }
}

$shares = Invoke-Collector 'SMB Shares' {
    if(-not(Get-Command Get-SmbShare -ErrorAction SilentlyContinue)){throw 'Get-SmbShare is unavailable.'}
    Get-SmbShare | Sort-Object Name | ForEach-Object {
        $share=$_
        $shareAcl=try{Get-SmbShareAccess $share.Name|ForEach-Object{"$($_.AccountName): $($_.AccessRight) ($($_.AccessControlType))"}}catch{@()}
        $ntfsAcl=try{if($share.Path -and (Test-Path -LiteralPath $share.Path)){(Get-Acl -LiteralPath $share.Path).Access|ForEach-Object{"$($_.IdentityReference): $($_.FileSystemRights) ($($_.AccessControlType))"}}}catch{@()}
        [pscustomobject][ordered]@{
            ComputerName=$ComputerName; Name=$share.Name; Path=$share.Path
            UNCPath=('\\{0}\{1}' -f $ComputerName,$share.Name); Description=$share.Description; Special=$share.Special
            SharePermissions=Join-Unique $shareAcl; NTFSPermissions=Join-Unique $ntfsAcl
        }
    }
}

$printers = Invoke-Collector 'Printers' {
    if(-not(Get-Command Get-Printer -ErrorAction SilentlyContinue)){throw 'Get-Printer is unavailable.'}
    Get-Printer | Sort-Object Name | Select-Object @{n='ComputerName';e={$ComputerName}},Name,ShareName,Shared,DriverName,PortName,Published
}

$accounts = Invoke-Collector 'Local Accounts' {
    if($features.Name -contains 'AD-Domain-Services'){ return }
    if(-not(Get-Command Get-LocalUser -ErrorAction SilentlyContinue)){throw 'Get-LocalUser is unavailable.'}
    Get-LocalUser | Sort-Object Name | Select-Object @{n='ComputerName';e={$ComputerName}},Name,Enabled,Description,LastLogon,PasswordExpires
}

$groups = Invoke-Collector 'Local Groups' {
    if($features.Name -contains 'AD-Domain-Services'){ return }
    if(-not(Get-Command Get-LocalGroupMember -ErrorAction SilentlyContinue)){throw 'Local group cmdlets are unavailable.'}
    foreach($group in Get-LocalGroup) {
        try {
            Get-LocalGroupMember $group.Name | Select-Object @{n='ComputerName';e={$ComputerName}},@{n='Group';e={$group.Name}},Name,ObjectClass,PrincipalSource
        } catch { Add-Diagnostic 'Local Groups' Warning ('{0}: {1}' -f $group.Name,$_.Exception.Message) }
    }
}

$tasks = Invoke-Collector 'Scheduled Tasks' {
    if(-not(Get-Command Get-ScheduledTask -ErrorAction SilentlyContinue)){throw 'Get-ScheduledTask is unavailable.'}
    Get-ScheduledTask | Where-Object {$IncludeMicrosoftTasks -or $_.TaskPath -notlike '\Microsoft\*'} |
      Sort-Object TaskPath,TaskName | ForEach-Object {
        [pscustomobject][ordered]@{
            Name=$_.TaskName; Server=$ComputerName; Location=('Task Scheduler: {0}' -f $_.TaskPath)
            Executable=Join-Unique ($_.Actions|ForEach-Object{('{0} {1}' -f $_.Execute,$_.Arguments).Trim()})
            State=$_.State; RunAsUser=$_.Principal.UserId; Description=$_.Description
        }
      }
}

$dc=@()
if($features.Name -contains 'AD-Domain-Services'){
    $dc=Invoke-Collector 'Domain Controller' {
        if(-not(Get-Module -ListAvailable ActiveDirectory)){throw 'ActiveDirectory module is unavailable.'}
        Import-Module ActiveDirectory
        $d=Get-ADDomainController -Identity $ComputerName; $domain=Get-ADDomain; $forest=Get-ADForest
        [pscustomobject][ordered]@{
            ComputerName=$ComputerName; FQDN=$d.HostName; Site=$d.Site; IPv4=$d.IPv4Address
            OperatingSystem=$d.OperatingSystem; GlobalCatalog=$d.IsGlobalCatalog; ReadOnly=$d.IsReadOnly
            DomainFQDN=$domain.DNSRoot; DomainShortName=$domain.NetBIOSName; UPNSuffixes=Join-Unique @($forest.UPNSuffixes)
            DomainMode=$domain.DomainMode; ForestMode=$forest.ForestMode; PDC=$domain.PDCEmulator
            RIDMaster=$domain.RIDMaster; InfrastructureMaster=$domain.InfrastructureMaster
            SchemaMaster=$forest.SchemaMaster; NamingMaster=$forest.DomainNamingMaster
        }
    }
}else{Add-Diagnostic 'Domain Controller' Success 'Skipped; AD DS role is not installed.'}

$adUsers=@()
$adMembership=@()
$adReplication=@()
if($features.Name -contains 'AD-Domain-Services'){
    $adUsers=Invoke-Collector 'AD Users' {
        Import-Module ActiveDirectory -ErrorAction Stop
        Get-ADUser -Filter * -Properties GivenName,Surname,DisplayName,UserPrincipalName,StreetAddress,City,State,PostalCode,co,Title,Department,Company,Manager,Description,Office,OfficePhone,EmailAddress,MobilePhone,Info,Enabled,LastLogonDate,ServicePrincipalName,PasswordNeverExpires |
          Sort-Object SamAccountName | ForEach-Object {
            [pscustomobject][ordered]@{
                FirstName=$_.GivenName; LastName=$_.Surname; DisplayName=$_.DisplayName
                SamAccountName=$_.SamAccountName; UserPrincipalName=$_.UserPrincipalName
                Street=$_.StreetAddress; City=$_.City; State=$_.State; PostalCode=$_.PostalCode; Country=$_.co
                JobTitle=$_.Title; Department=$_.Department; Company=$_.Company; Manager=$_.Manager
                Description=$_.Description; Office=$_.Office; Telephone=$_.OfficePhone; Email=$_.EmailAddress
                Mobile=$_.MobilePhone; Notes=$_.Info; Enabled=$_.Enabled; LastLogonDate=$_.LastLogonDate
                LikelyServiceAccount=($_.SamAccountName -match '\$' -or @($_.ServicePrincipalName).Count -gt 0)
                PasswordNeverExpires=$_.PasswordNeverExpires
            }
          }
    }

    $adMembership=Invoke-Collector 'AD Group Membership' {
        Import-Module ActiveDirectory -ErrorAction Stop
        $domain=Get-ADDomain
        $primaryGroups=@{}
        foreach($user in Get-ADUser -Filter * -Properties DisplayName,MemberOf,PrimaryGroupID){
            foreach($groupDn in @($user.MemberOf)){
                $groupName=(($groupDn -split '(?<!\\),')[0] -replace '^CN=','') -replace '\\,',','
                [pscustomobject][ordered]@{
                    Username=$user.SamAccountName; Name=$user.DisplayName; GroupName=$groupName; MembershipType='Direct'
                }
            }
            $rid=[string]$user.PrimaryGroupID
            if($rid){
                if(-not $primaryGroups.ContainsKey($rid)){
                    $primarySid='{0}-{1}' -f $domain.DomainSID.Value,$rid
                    $primaryGroups[$rid]=(Get-ADGroup -Identity $primarySid).Name
                }
                [pscustomobject][ordered]@{
                    Username=$user.SamAccountName; Name=$user.DisplayName; GroupName=$primaryGroups[$rid]; MembershipType='Primary'
                }
            }
        }
    }

    $adReplication=Invoke-Collector 'AD Replication' {
        Import-Module ActiveDirectory -ErrorAction Stop
        Get-ADReplicationPartnerMetadata -Target $ComputerName -Scope Server | ForEach-Object {
            [pscustomobject][ordered]@{
                FromServer=$_.Partner; ToServer=$_.Server; LastSync=$_.LastReplicationSuccess
                Status=$(if($_.LastReplicationResult -eq 0){'Success'}else{'Error {0}' -f $_.LastReplicationResult})
            }
        }
    }
}else{
    Add-Diagnostic 'AD Users' Success 'Skipped; AD DS role is not installed.'
    Add-Diagnostic 'AD Group Membership' Success 'Skipped; AD DS role is not installed.'
    Add-Diagnostic 'AD Replication' Success 'Skipped; AD DS role is not installed.'
}

$dhcp=@()
if($features.Name -contains 'DHCP'){
    $dhcp=Invoke-Collector 'DHCP Scopes' {
        if(-not(Get-Command Get-DhcpServerv4Scope -ErrorAction SilentlyContinue)){throw 'DHCP cmdlets are unavailable.'}
        foreach($s in Get-DhcpServerv4Scope -ComputerName $ComputerName){
            $stats=try{Get-DhcpServerv4ScopeStatistics -ComputerName $ComputerName -ScopeId $s.ScopeId}catch{$null}
            $exclusions=try{Get-DhcpServerv4ExclusionRange -ComputerName $ComputerName -ScopeId $s.ScopeId}catch{@()}
            $reservations=try{Get-DhcpServerv4Reservation -ComputerName $ComputerName -ScopeId $s.ScopeId}catch{@()}
            $options=try{Get-DhcpServerv4OptionValue -ComputerName $ComputerName -ScopeId $s.ScopeId}catch{@()}
            $failover=try{Get-DhcpServerv4Failover -ComputerName $ComputerName -ScopeId $s.ScopeId}catch{$null}
            [pscustomobject][ordered]@{
                ServerName=$ComputerName; ScopeName=$s.Name; Status=$s.State; ScopeId=$s.ScopeId
                StartRange=$s.StartRange; EndRange=$s.EndRange; SubnetMask=$s.SubnetMask
                AddressPool=('{0} - {1}' -f $s.StartRange,$s.EndRange)
                Exclusions=Join-Unique @($exclusions|ForEach-Object{'{0} - {1}' -f $_.StartRange,$_.EndRange})
                Reservations=Join-Unique @($reservations|ForEach-Object{'{0} ({1})' -f $_.IPAddress,$_.Name})
                ScopeOptions=@($options|ForEach-Object{'{0} {1}: {2}' -f $_.OptionId,$_.Name,($_.Value -join ', ')})
                LeaseDuration=$s.LeaseDuration; Utilization=$(if($stats){$stats.PercentageInUse}else{$null})
                NAP=$(if($s.PSObject.Properties['NapEnable']){$s.NapEnable}else{$null})
                Failover=$(if($failover){$failover.Name}else{$null})
            }
        }
    }
}else{Add-Diagnostic 'DHCP Scopes' Success 'Skipped; DHCP role is not installed.'}

$dns=@()
if($features.Name -contains 'DNS'){
    $dns=Invoke-Collector 'DNS Zones' {
        if(-not(Get-Command Get-DnsServerZone -ErrorAction SilentlyContinue)){throw 'DNS cmdlets are unavailable.'}
        Get-DnsServerZone -ComputerName $ComputerName | Sort-Object ZoneName |
          Select-Object @{n='ComputerName';e={$ComputerName}},ZoneName,ZoneType,IsReverseLookupZone,IsDsIntegrated,DynamicUpdate,ReplicationScope
    }
}else{Add-Diagnostic 'DNS Zones' Success 'Skipped; DNS role is not installed.'}

$dhcpClient=Invoke-Collector 'DHCP Client Configuration' {
    Get-CimInstance Win32_NetworkAdapterConfiguration -Filter 'IPEnabled=True' |
      Where-Object DHCPEnabled | Select-Object -ExpandProperty DHCPServer
}
$mapped=Invoke-Collector 'Mapped Drives' {
    Get-CimInstance Win32_LogicalDisk -Filter 'DriveType=4' | ForEach-Object {'{0}={1}' -f $_.DeviceID,$_.ProviderName}
}

$serviceAccounts=@(
    $services | Where-Object {
        $_.StartName -and $_.StartName -notmatch '^(LocalSystem|NT AUTHORITY\\|NT SERVICE\\|LocalService|NetworkService)
try {
    $package=Open-ExcelPackage -Path $outputPath
    $book=$package.Workbook
    $sys=$system|Select-Object -First 1
    $nics=@($network|Select-Object -First 2)
    $serverValues=@{
        'Name'=$ComputerName; 'Physical/Virtual'=$sys.PhysicalOrVirtual
        'OS'=('{0} ({1}, build {2})' -f $sys.OperatingSystem,$sys.Version,$sys.Build)
        'CPU/vCPU'=$sys.LogicalProcessors; '# of CPUs'=$sys.Sockets; '# of Cores'=$sys.Cores
        'RAM/vRAM'=('{0} GB' -f $sys.MemoryGB); 'Open Ports'=Join-Unique $ports.LocalPort ', '
        'DHCP Server'=Join-Unique $dhcpClient; 'Mapped Drives'=Join-Unique $mapped
    }
    for($i=0;$i -lt [math]::Min($disks.Count,7);$i++){
        $h=if($i -eq 0){'HDD1 Free Space'}else{'HDD{0}' -f ($i+1)}
        $serverValues[$h]='{0} {1} GB free / {2} GB' -f $disks[$i].Drive,$disks[$i].FreeGB,$disks[$i].SizeGB
    }
    for($i=0;$i -lt $nics.Count;$i++){
        $number=$i+1; $dnsValues=@($nics[$i].DNS -split ';\s*')
        $serverValues["NIC$number IP"]=$nics[$i].IPAddress
        $serverValues["NIC$number Subnet"]=$nics[$i].SubnetMask
        $serverValues["NIC$number Gateway"]=$nics[$i].Gateway
        if($number -eq 1){
            for($d=0;$d -lt [math]::Min($dnsValues.Count,3);$d++){$serverValues["NIC1 DNS$($d+1)"]=$dnsValues[$d]}
        }else{
            if($dnsValues.Count){$serverValues['NIC2 DNS']=$dnsValues[0]}
            if($dnsValues.Count -gt 1){$serverValues['NIC2 DNS2']=$dnsValues[1]}
        }
    }
    Set-TemplateRow $book.Worksheets['Server Info'] 1 2 $serverValues

    $roleSheet=$book.Worksheets['Server Roles']
    $roleMap=[ordered]@{
      'ADCS'='AD-Certificate';'ADDS'='AD-Domain-Services';'ADFS'='ADFS-Federation';'ADLDS'='ADLDS';'ADRMS'='ADRMS'
      'Application Server'='Application-Server';'DHCP'='DHCP';'DNS'='DNS';'Fax Server'='Fax';'File Server'='FS-FileServer'
      'File Resource Manager'='FS-Resource-Manager';'Data Deduplication'='FS-Data-Deduplication';'DFS Namespaces'='FS-DFS-Namespace'
      'DFS Replication'='FS-DFS-Replication';'iSCSI Target Server'='FS-iSCSITarget-Server';'Storage Services'='Storage-Services'
      'Hyper-V'='Hyper-V';'NPS'='NPAS';'Print Server'='Print-Server';'Distributed Scan'='Print-Scan-Server'
      'Remote Access'='RemoteAccess';'RDS'='Remote-Desktop-Services';'RDS Connection Broker'='RDS-Connection-Broker'
      'RDS Gateway'='RDS-Gateway';'RDS Licensing'='RDS-Licensing';'RDS Session Host'='RDS-RD-Server'
      'RDS Virtualization Host'='RDS-Virtualization';'RDS Web Access'='RDS-Web-Access';'Volume Activation Services'='VolumeActivation'
      'Web Server (IIS)'='Web-Server';'Web Server Management Tools'='Web-Mgmt-Tools';'WDS'='WDS';'WSUS'='UpdateServices'
      'Windows Server Essentials Experience'='ServerEssentialsRole'
    }
    $rv=New-YesNoTemplateRow $roleSheet 1
    foreach($e in $roleMap.GetEnumerator()){$rv[$e.Key]=if($features.Name -contains $e.Value){'Yes'}else{'No'}}
    Set-TemplateRow $roleSheet 1 2 $rv

    $featureSheet=$book.Worksheets['Server Features']
    $featureMap=[ordered]@{
      '.NET Framework 3.5 Features'='NET-Framework-Features';'.NET Framework 4.5/4.6/4.7 Features'='NET-Framework-45-Features'
      'ASP.NET 4.7'='NET-Framework-45-ASPNET';'Web Server Tools'='Web-Mgmt-Tools';'Windows Communication Foundation'='NET-WCF-Services45'
      'BITS'='BITS';'BitLocker Drive Encryption'='BitLocker';'Enhanced Storage'='EnhancedStorage';'Failover Clustering'='Failover-Clustering'
      'Group Policy Management'='GPMC';'I/O Quality of Service'='DiskIo-QoS';'IIS Hostable Web Core'='Web-WHC'
      'Ink and Handwriting Services'='InkAndHandwritingServices';'IPAM'='IPAM';'Media Foundation'='Server-Media-Foundation'
      'Remote Assistance'='Remote-Assistance';'Message Queuing'='MSMQ';'Remote Differential Compression'='RDC'
      'Multipath I/O'='Multipath-IO';'Network Load Balancing'='NLB';'RSAT'='RSAT';'RPC over HTTP Proxy'='RPC-over-HTTP-Proxy'
      'RDS Tools'='RSAT-RDS-Tools';'SMB 1.0 / CIFS File Sharing Support'='FS-SMB1';'SMTP Server'='SMTP-Server'
      'SNMP Service'='SNMP-Service';'System Data Archiver'='System-DataArchiver';'Telnet Client'='Telnet-Client';'TFTP Client'='TFTP-Client'
      'Windows Defender Features'='Windows-Defender';'Windows Identity Foundation 3.5'='Windows-Identity-Foundation'
      'Windows Internal Database'='Windows-Internal-Database';'Windows PowerShell'='PowerShellRoot'
      'Windows Process Activation Service'='WAS';'Windows Search Services'='Search-Service';'Windows Server Backup'='Windows-Server-Backup'
      'Windows Server Migration Tools'='Migration';'WINS Server'='WINS';'WoW64 Support'='WoW64-Support';'XPS Viewer'='XPS-Viewer'
    }
    $fv=New-YesNoTemplateRow $featureSheet 1
    foreach($e in $featureMap.GetEnumerator()){$fv[$e.Key]=if($features.Name -contains $e.Value){'Yes'}else{'No'}}
    Set-TemplateRow $featureSheet 1 2 $fv

    $software=Join-Unique @($applications.Name+$services.Name+$services.DisplayName) ([Environment]::NewLine)
    $azureSheet=$book.Worksheets['Azure Agents-Apps-Services']
    $azurePatterns=[ordered]@{
      'MMA'='Microsoft Monitoring Agent|OMS Agent';'Azure Data Studio'='Azure Data Studio'
      'Azure Workload Backup'='Azure.*Workload|Microsoft Azure Recovery Services';'Azure Powershell'='Azure PowerShell|Az PowerShell'
      'UI Flow'='Power Automate|UI flows';'On-Premises Data Gateway'='On-premises data gateway'
      'Azure Plug-in for Veeam'='Veeam.*Azure|Azure.*Veeam';'Azure Backup'='Azure Backup'
      'Azure Site Recovery Mobility Service'='Site Recovery.*Mobility|Microsoft Azure Site Recovery'
      'Windows Azure VM Agent'='Windows Azure VM Agent|WindowsAzureGuestAgent|RdAgent'
      'Azure AD Connect Authentication Agent'='Azure AD Connect Authentication Agent|Entra Connect Authentication Agent'
      'Azure AD Connect'='Azure AD Connect|Microsoft Entra Connect';'Citrix Azure Provisioning'='Citrix.*Azure'
      'Azure Connected Storage Service'='Azure Connected Storage';'Azure Recovery Services'='Azure Recovery Services'
      'Azure Information Protection'='Azure Information Protection';'Azure Advanced Threat Protection Sensor'='Azure ATP|Defender for Identity'
      'Azure AD Connect Agent Updater'='Azure AD Connect Agent Updater|Entra Connect Agent Updater'
      'NPS Extension for Azure MFA'='NPS Extension.*Azure MFA|Azure MFA.*NPS';'Intune Connector dor AD'='Intune Connector.*Active Directory'
      'Microsoft Azure Active Directory Application Proxy Connector'='Application Proxy Connector'
      'Storage Sync Agent'='Storage Sync Agent|Azure File Sync'
    }
    $av=New-YesNoTemplateRow $azureSheet 1
    $av['Azure VM']=if($sys.Manufacturer -match 'Microsoft' -and $sys.Model -match 'Virtual'){'Yes'}else{'No'}
    foreach($e in $azurePatterns.GetEnumerator()){$av[$e.Key]=if($software -match $e.Value){'Yes'}else{'No'}}
    Set-TemplateRow $azureSheet 1 2 $av

    $lob=$book.Worksheets['LoB Applications']
    if($lob){
        Clear-TemplateValues $lob 'B2:BO2'
        $lob.Cells[2,1].Value=$ComputerName
        for($i=0;$i -lt [math]::Min($applications.Count,$lob.Dimension.End.Column-1);$i++){
            $lob.Cells[2,($i+2)].Value=$applications[$i].Name
        }
    }

    $shareSheet=$book.Worksheets['File Server Shares']
    $securitySheet=$book.Worksheets['File Server Security']
    Clear-TemplateValues $shareSheet 'A3:G65'
    Clear-TemplateValues $securitySheet 'A2:G39'
    for($i=0;$i -lt $shares.Count;$i++){
        if($shareSheet){
            $r=$i+3;$shareSheet.Cells[$r,1].Value=$ComputerName;$shareSheet.Cells[$r,2].Value=$shares[$i].Path
            $shareSheet.Cells[$r,3].Value=$(if($shares[$i].Path -match '^[A-Za-z]:'){$shares[$i].Path.Substring(0,2)}else{$null})
            $shareSheet.Cells[$r,5].Value=$shares[$i].Name;$shareSheet.Cells[$r,6].Value=$shares[$i].UNCPath;$shareSheet.Cells[$r,7].Value=$shares[$i].Path
        }
        if($securitySheet){
            $r=$i+2;$securitySheet.Cells[$r,1].Value=$ComputerName;$securitySheet.Cells[$r,2].Value=$shares[$i].Path
            $securitySheet.Cells[$r,3].Value=$shares[$i].UNCPath;$securitySheet.Cells[$r,4].Value=$shares[$i].SharePermissions
            $securitySheet.Cells[$r,6].Value=$shares[$i].NTFSPermissions
        }
    }

    $printerSheet=$book.Worksheets['Printers Shares']
    if($printerSheet){
        Clear-TemplateValues $printerSheet 'A3:B100'
        $printerSheet.Cells[1,2].Value=$ComputerName
        for($i=0;$i -lt $printers.Count;$i++){
            $r=$i+3
            $printerSheet.Cells[$r,1].Value=$(if($printers[$i].ShareName){$printers[$i].ShareName}else{$printers[$i].Name})
            $printerSheet.Cells[$r,2].Value=$printers[$i].PortName
        }
    }

    $taskSheet=$book.Worksheets['Scheduled Tasks']
    Clear-TemplateValues $taskSheet 'A2:E38'
    for($i=0;$taskSheet -and $i -lt $tasks.Count;$i++){
        $r=$i+2;$taskSheet.Cells[$r,1].Value=$tasks[$i].Name;$taskSheet.Cells[$r,2].Value=$ComputerName
        $taskSheet.Cells[$r,3].Value=$tasks[$i].Location;$taskSheet.Cells[$r,4].Value=$tasks[$i].Executable
        $taskSheet.Cells[$r,5].Value=$tasks[$i].Description
    }

    $accountSheet=$book.Worksheets['Accounts and Groups']
    if($accountSheet){
        Clear-TemplateValues $accountSheet 'A2:B4'
        Clear-TemplateValues $accountSheet 'A8:B10'
        Clear-TemplateValues $accountSheet 'A14:D16'
        for($i=0;$i -lt [math]::Min($accounts.Count,3);$i++){
            $accountSheet.Cells[$i+2,1].Value=$accounts[$i].Name
        }
        if($adUsers.Count){
            $accountSheet.Cells[8,1].Value=('{0} domain accounts; see User Accounts' -f $adUsers.Count)
            $accountSheet.Cells[8,2].Value='See SecurityGroups for memberships'
        }
        for($i=0;$i -lt [math]::Min($serviceAccounts.Count,3);$i++){
            $accountSheet.Cells[$i+14,1].Value=$serviceAccounts[$i].Account
            $accountSheet.Cells[$i+14,2].Value=$serviceAccounts[$i].Usage
            $accountSheet.Cells[$i+14,3].Value=$serviceAccounts[$i].InteractiveLoginRequired
            $accountSheet.Cells[$i+14,4].Value=$serviceAccounts[$i].ManagedServiceAccount
        }
    }

    if($dc.Count){
        $ad=$book.Worksheets['Active Directory'];$d=$dc[0]
        if($ad){
            $ad.Cells[2,2].Value=('Domain FQDN: {0}' -f $d.DomainFQDN)
            $ad.Cells[2,4].Value=('Domain Shortname: {0}' -f $d.DomainShortName)
            $ad.Cells[2,6].Value=('Additional UPN: {0}' -f $d.UPNSuffixes)
            $ad.Cells[4,2].Value=$d.FQDN;$ad.Cells[4,3].Value=$d.OperatingSystem;$ad.Cells[4,4].Value='Online'
            $ad.Cells[4,5].Value=if($d.GlobalCatalog){'Yes'}else{'No'};$ad.Cells[4,6].Value=$d.Site
            $ad.Cells[4,7].Value=if($d.ReadOnly){'Yes'}else{'No'};$ad.Cells[4,8].Value=if($features.Name -contains 'DHCP'){'Yes'}else{'No'}
            $ad.Cells[13,2].Value=$d.FQDN;$ad.Cells[13,3].Value=if($d.PDC -eq $d.FQDN){'x'}else{''}
            $ad.Cells[13,4].Value=if($d.SchemaMaster -eq $d.FQDN){'x'}else{''};$ad.Cells[13,5].Value=if($d.NamingMaster -eq $d.FQDN){'x'}else{''}
            $ad.Cells[13,6].Value=if($d.RIDMaster -eq $d.FQDN){'x'}else{''};$ad.Cells[13,7].Value=if($d.InfrastructureMaster -eq $d.FQDN){'x'}else{''}
            $ad.Cells[21,2].Value=$d.ForestMode;$ad.Cells[21,3].Value=$d.DomainMode
            Clear-TemplateValues $ad 'B24:E38'
            for($i=0;$i -lt [math]::Min($adReplication.Count,15);$i++){
                $r=$i+24;$ad.Cells[$r,2].Value=$adReplication[$i].FromServer;$ad.Cells[$r,3].Value=$adReplication[$i].ToServer
                $ad.Cells[$r,4].Value=$adReplication[$i].LastSync;$ad.Cells[$r,4].Style.Numberformat.Format='yyyy-mm-dd HH:mm:ss'
                $ad.Cells[$r,5].Value=$adReplication[$i].Status
            }
            $ad.Cells[41,2].Value=$d.FQDN
            $ad.Cells[41,3].Value=if($services.Name -contains 'ADSync'){'Configured'}else{'Not detected'}
        }
    }

    $dhcpSheet=$book.Worksheets['DHCP']
    Clear-TemplateValues $dhcpSheet 'A4:P9'
    for($i=0;$dhcpSheet -and $i -lt $dhcp.Count;$i++){
        $s=$dhcp[$i];$optionValues=@($s.ScopeOptions)
        $vals=@{
          'Server Name'=$ComputerName;'Scope Name'=$s.ScopeName;'Status'=$s.Status
          'Scope'=('{0} / {1}' -f $s.ScopeId,$s.SubnetMask);'Address Pool'=$s.AddressPool
          'Exclusions'=$s.Exclusions;'Reservations'=$s.Reservations
          'Scope Utilization'=$s.Utilization;'Lease Duration'=$s.LeaseDuration;'NAP'=$s.NAP;'Failover'=$s.Failover
        }
        for($optionIndex=0;$optionIndex -lt [math]::Min($optionValues.Count,5);$optionIndex++){
            $vals['Scope Options{0}' -f ($optionIndex+1)]=$optionValues[$optionIndex]
        }
        Set-TemplateRow $dhcpSheet 3 ($i+4) $vals
    }

    $dnsSheet=$book.Worksheets['DNS']
    if($dnsSheet){
        Clear-TemplateValues $dnsSheet 'B4:C13'
        Clear-TemplateValues $dnsSheet 'E4:H13'
        Clear-TemplateValues $dnsSheet 'E16:H31'
        $dnsSheet.Cells[4,2].Value=$(if($dc.Count){$dc[0].FQDN}else{$ComputerName})
        $dnsSheet.Cells[4,3].Value='Online'
        $forward=@($dns|Where-Object{-not $_.IsReverseLookupZone})
        $reverse=@($dns|Where-Object{$_.IsReverseLookupZone})
        for($i=0;$i -lt [math]::Min($forward.Count,10);$i++){
            $r=$i+4;$dnsSheet.Cells[$r,5].Value=$forward[$i].ZoneName;$dnsSheet.Cells[$r,6].Value=$forward[$i].ZoneType
            $dnsSheet.Cells[$r,7].Value='Loaded';$dnsSheet.Cells[$r,8].Value='Not collected'
        }
        for($i=0;$i -lt [math]::Min($reverse.Count,16);$i++){
            $r=$i+16;$dnsSheet.Cells[$r,5].Value=$reverse[$i].ZoneName;$dnsSheet.Cells[$r,6].Value=$reverse[$i].ZoneType
            $dnsSheet.Cells[$r,7].Value='Loaded';$dnsSheet.Cells[$r,8].Value='Not collected'
        }
    }

    $userSheet=$book.Worksheets['User Accounts']
    if($userSheet){
        for($i=0;$i -lt $adUsers.Count;$i++){
            $u=$adUsers[$i];$r=$i+3
            $values=@($u.FirstName,$u.LastName,$u.DisplayName,$u.SamAccountName,$u.UserPrincipalName,$u.Street,$u.City,$u.State,$u.PostalCode,$u.Country,
              $u.JobTitle,$u.Department,$u.Company,$u.Manager,$u.Description,$u.Office,$u.Telephone,$u.Email,$u.Mobile,$u.Notes,
              $(if($u.Enabled){'Enabled'}else{'Disabled'}),$u.LastLogonDate)
            for($c=0;$c -lt $values.Count;$c++){$userSheet.Cells[$r,($c+1)].Value=$values[$c]}
            if($u.LastLogonDate){$userSheet.Cells[$r,22].Style.Numberformat.Format='yyyy-mm-dd HH:mm:ss'}
        }
    }

    $groupSheet=$book.Worksheets['SecurityGroups']
    if($groupSheet){
        for($i=0;$i -lt $adMembership.Count;$i++){
            $r=$i+3;$groupSheet.Cells[$r,1].Value=$adMembership[$i].Username
            $groupSheet.Cells[$r,2].Value=$adMembership[$i].Name;$groupSheet.Cells[$r,3].Value=$adMembership[$i].GroupName
        }
    }

    Add-DataSheet $book 'System Inventory' $system
    Add-DataSheet $book 'Network Inventory' $network
    Add-DataSheet $book 'Disk Inventory' $disks
    Add-DataSheet $book 'Listening Ports' $ports
    Add-DataSheet $book 'Application Inventory' $applications
    Add-DataSheet $book 'Services' $services
    Add-DataSheet $book 'Installed Features' $features
    Add-DataSheet $book 'SMB Share Inventory' $shares
    Add-DataSheet $book 'Printer Inventory' $printers
    Add-DataSheet $book 'Local Accounts' $accounts
    Add-DataSheet $book 'Local Group Membership' $groups
    Add-DataSheet $book 'Task Inventory' $tasks
    Add-DataSheet $book 'Domain Controller Inventory' $dc
    Add-DataSheet $book 'DHCP Scope Inventory' $dhcp
    Add-DataSheet $book 'DNS Zone Inventory' $dns
    Add-DataSheet $book 'Diagnostics' $script:Diagnostics
    Close-ExcelPackage -ExcelPackage $package
    $package=$null
} catch {
    if($package){$package.Dispose()}
    throw
}

Write-Host ''
Write-Host 'Server discovery completed.' -ForegroundColor Green
Write-Host ('Output: {0}' -f $outputPath)
Write-Host ('Collector failures: {0}' -f @($script:Diagnostics|Where-Object Status -eq Failed).Count)
Write-Output $outputPath
