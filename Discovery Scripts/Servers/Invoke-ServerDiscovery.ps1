#Requires -Version 5.1
<#
.SYNOPSIS
Collects discovery data from the local Windows Server and writes a clean,
timestamped Excel discovery workbook.

.DESCRIPTION
Phase 1 is local-only. It does not enumerate AD computers, use remoting, call
Win32_Product, or recursively scan file shares. It installs ImportExcel when
required. Each collector is isolated; failures are recorded on Diagnostics.
The workbook is generated from the collected data; a legacy template is optional.
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
    $joined=(($Values | Where-Object { $null -ne $_ -and "$_".Trim() } |
        ForEach-Object { "$_".Trim() } | Sort-Object -Unique) -join $Separator)
    if ([string]::IsNullOrWhiteSpace($joined)) { return $null }
    return $joined
}

function Convert-PrefixToMask {
    param([ValidateRange(0,32)][int]$Prefix)
    $bits = (('1' * $Prefix) -join '').PadRight(32,'0')
    return ((0,8,16,24 | ForEach-Object {
        [Convert]::ToInt32($bits.Substring($_,8),2)
    }) -join '.')
}

function Convert-ToCellValue {
    param($Value)
    if ($null -eq $Value) { return $null }
    if ($Value -is [bool]) { return $(if ($Value) { 'Yes' } else { 'No' }) }
    if ($Value -is [array]) { return (Join-Unique $Value) }
    if ($Value -is [datetime] -or $Value -is [string] -or $Value -is [decimal] -or $Value.GetType().IsPrimitive) {
        return $Value
    }
    return [string]$Value
}

function Set-WorksheetStyle {
    param($Worksheet,[int]$HeaderRow = 1)
    if (-not $Worksheet -or -not $Worksheet.Dimension) { return }

    $lastRow=$Worksheet.Dimension.End.Row
    $lastColumn=$Worksheet.Dimension.End.Column
    $header=$Worksheet.Cells[$HeaderRow,1,$HeaderRow,$lastColumn]
    $header.Style.Font.Bold=$true
    $header.Style.Font.Color.SetColor([System.Drawing.Color]::White)
    $header.Style.Fill.PatternType=[OfficeOpenXml.Style.ExcelFillStyle]::Solid
    $header.Style.Fill.BackgroundColor.SetColor([System.Drawing.Color]::FromArgb(31,78,121))
    $header.Style.HorizontalAlignment=[OfficeOpenXml.Style.ExcelHorizontalAlignment]::Center
    $header.Style.VerticalAlignment=[OfficeOpenXml.Style.ExcelVerticalAlignment]::Center
    $header.Style.WrapText=$true
    $header.AutoFilter=$true

    if ($lastRow -gt $HeaderRow) {
        $body=$Worksheet.Cells[($HeaderRow+1),1,$lastRow,$lastColumn]
        $body.Style.VerticalAlignment=[OfficeOpenXml.Style.ExcelVerticalAlignment]::Top
    }

    $Worksheet.Cells[$Worksheet.Dimension.Address].Style.Font.Name='Arial'
    $Worksheet.Cells[$Worksheet.Dimension.Address].Style.Font.Size=10
    $Worksheet.Cells[$Worksheet.Dimension.Address].AutoFitColumns()
    for ($column=1; $column -le $lastColumn; $column++) {
        if ($Worksheet.Column($column).Width -lt 12) { $Worksheet.Column($column).Width=12 }
        if ($Worksheet.Column($column).Width -gt 45) {
            $Worksheet.Column($column).Width=45
            if ($lastRow -gt $HeaderRow) {
                $Worksheet.Cells[($HeaderRow+1),$column,$lastRow,$column].Style.WrapText=$true
            }
        }
    }
    $Worksheet.View.ShowGridLines=$false
    $Worksheet.View.FreezePanes(($HeaderRow+1),1)
}

function Add-DataSheet {
    param($Workbook,[string]$Name,[object[]]$Rows)
    if ($Workbook.Worksheets[$Name]) { $Workbook.Worksheets.Delete($Name) }
    $sheet=$Workbook.Worksheets.Add($Name)
    $data=@($Rows)

    if (-not $data.Count) {
        $sheet.Cells[1,1].Value='Status'
        $sheet.Cells[2,1].Value='No records returned.'
        Set-WorksheetStyle $sheet
        return $sheet
    }

    $headers=@($data[0].PSObject.Properties.Name)
    for ($column=0; $column -lt $headers.Count; $column++) {
        $sheet.Cells[1,($column+1)].Value=$headers[$column]
    }

    for ($row=0; $row -lt $data.Count; $row++) {
        for ($column=0; $column -lt $headers.Count; $column++) {
            $propertyName=$headers[$column]
            $value=Convert-ToCellValue $data[$row].PSObject.Properties[$propertyName].Value
            $cell=$sheet.Cells[($row+2),($column+1)]
            $cell.Value=$value
            if ($value -is [datetime]) {
                $cell.Style.Numberformat.Format='yyyy-mm-dd HH:mm:ss'
            } elseif ($propertyName -match 'Percent|Utilization|SizeGB|FreeGB|MemoryGB') {
                $cell.Style.Numberformat.Format='0.00'
            }
        }
    }

    Set-WorksheetStyle $sheet
    return $sheet
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
$template=$null
if (-not [string]::IsNullOrWhiteSpace($TemplatePath)) {
    $template=Get-Item -LiteralPath $TemplatePath
    if ($template.Extension -ne '.xlsx') { throw 'TemplatePath must be an .xlsx file.' }
}
if (-not (Test-Path -LiteralPath $OutputDirectory)) {
    New-Item -ItemType Directory -Path $OutputDirectory -Force | Out-Null
}
$outputPath=Join-Path $OutputDirectory ('{0}-ServerDiscovery-{1}.xlsx' -f $ComputerName,(Get-Date -Format yyyyMMdd-HHmmss))
if ($template) {
    Copy-Item -LiteralPath $template.FullName -Destination $outputPath -Force
}

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
                LikelyServiceAccount=($_.SamAccountName.EndsWith('$') -or @($_.ServicePrincipalName).Count -gt 0)
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
$dhcpExclusions=@()
$dhcpReservations=@()
$dhcpOptions=@()
$dhcpFailover=@()
if($features.Name -contains 'DHCP'){
    if(-not(Get-Command Get-DhcpServerv4Scope -ErrorAction SilentlyContinue)){throw 'DHCP cmdlets are unavailable.'}

    $dhcp=Invoke-Collector 'DHCP Scopes' {
        foreach($s in Get-DhcpServerv4Scope -ComputerName $ComputerName -ErrorAction Stop){
            $stats=try{
                Get-DhcpServerv4ScopeStatistics -ComputerName $ComputerName -ScopeId $s.ScopeId -ErrorAction Stop
            }catch{
                Add-Diagnostic 'DHCP Scope Statistics' Warning ('{0}: {1}' -f $s.ScopeId,$_.Exception.Message)
                $null
            }
            [pscustomobject][ordered]@{
                ServerName=$ComputerName
                ScopeName=$s.Name
                Status=$s.State
                ScopeId=$s.ScopeId
                StartRange=$s.StartRange
                EndRange=$s.EndRange
                SubnetMask=$s.SubnetMask
                AddressPool=('{0} - {1}' -f $s.StartRange,$s.EndRange)
                LeaseDuration=$s.LeaseDuration
                Utilization=$(if($stats){$stats.PercentageInUse}else{$null})
                NAP=$(if($s.PSObject.Properties['NapEnable']){$s.NapEnable}else{$null})
            }
        }
    }

    $dhcpExclusions=Invoke-Collector 'DHCP Exclusions' {
        foreach($s in $dhcp){
            Get-DhcpServerv4ExclusionRange -ComputerName $ComputerName -ScopeId $s.ScopeId -ErrorAction Stop |
              ForEach-Object {
                [pscustomobject][ordered]@{
                    ServerName=$ComputerName
                    ScopeId=$s.ScopeId
                    ScopeName=$s.ScopeName
                    StartRange=$_.StartRange
                    EndRange=$_.EndRange
                }
              }
        }
    }

    $dhcpReservations=Invoke-Collector 'DHCP Reservations' {
        foreach($s in $dhcp){
            Get-DhcpServerv4Reservation -ComputerName $ComputerName -ScopeId $s.ScopeId -ErrorAction Stop |
              ForEach-Object {
                [pscustomobject][ordered]@{
                    ServerName=$ComputerName
                    ScopeId=$s.ScopeId
                    ScopeName=$s.ScopeName
                    IPAddress=$_.IPAddress
                    ClientId=$_.ClientId
                    Name=$_.Name
                    Description=$_.Description
                    Type=$_.Type
                }
              }
        }
    }

    $dhcpOptions=Invoke-Collector 'DHCP Options' {
        foreach($s in $dhcp){
            Get-DhcpServerv4OptionValue -ComputerName $ComputerName -ScopeId $s.ScopeId -ErrorAction Stop |
              ForEach-Object {
                [pscustomobject][ordered]@{
                    ServerName=$ComputerName
                    ScopeId=$s.ScopeId
                    ScopeName=$s.ScopeName
                    OptionId=$_.OptionId
                    Name=$_.Name
                    Type=$_.Type
                    Value=(Join-Unique @($_.Value) ', ')
                    VendorClass=$_.VendorClass
                    UserClass=$_.UserClass
                    PolicyName=$_.PolicyName
                }
              }
        }
    }

    $dhcpFailover=Invoke-Collector 'DHCP Failover' {
        Get-DhcpServerv4Failover -ComputerName $ComputerName -ErrorAction SilentlyContinue |
          ForEach-Object {
            [pscustomobject][ordered]@{
                ServerName=$ComputerName
                Name=$_.Name
                PartnerServer=$_.PartnerServer
                Mode=$_.Mode
                State=$_.State
                ServerRole=$_.ServerRole
                ScopeId=(Join-Unique @($_.ScopeId))
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
        Add-Diagnostic $collectorName Success 'Skipped; DHCP role is not installed.'
    }
}

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
        $_.StartName -and $_.StartName -notmatch '^(LocalSystem|LocalService|NetworkService|NT AUTHORITY\\|NT SERVICE\\)'
    } | ForEach-Object {
        [pscustomobject][ordered]@{
            Account=$_.StartName
            Usage=('Windows service: {0}' -f $_.DisplayName)
            InteractiveLoginRequired='Unknown'
            ManagedServiceAccount=$(if($_.StartName.EndsWith('$')){'Yes'}else{'No'})
        }
    }
)
$serviceAccounts=@($serviceAccounts | Sort-Object Account,Usage -Unique)

$package=$null
try {
    if ($template) {
        $package=Open-ExcelPackage -Path $outputPath
    } else {
        $package=Open-ExcelPackage -Path $outputPath -Create
    }
    $book=$package.Workbook

    # Rebuild the workbook around normalized data. If a legacy template was
    # supplied, its worksheets are intentionally replaced to avoid duplicate
    # summary/detail tabs and stale sample values.
    $existingSheetNames=@($book.Worksheets | ForEach-Object Name)
    $workingSheetName='__DiscoveryBuild'
    while ($book.Worksheets[$workingSheetName]) { $workingSheetName='_' + $workingSheetName }
    $null=$book.Worksheets.Add($workingSheetName)
    foreach ($sheetName in $existingSheetNames) { $book.Worksheets.Delete($sheetName) }

    $sys=$system | Select-Object -First 1
    $installedRoleCount=@($features | Where-Object FeatureType -eq Role).Count
    $installedFeatureCount=@($features | Where-Object FeatureType -ne Role).Count
    $ipv4Addresses=Join-Unique $network.IPAddress
    $gateways=Join-Unique $network.Gateway
    $dnsServers=Join-Unique $network.DNS
    $diskTotal=($disks | Measure-Object SizeGB -Sum).Sum
    $diskFree=($disks | Measure-Object FreeGB -Sum).Sum
    $collectorFailures=@($script:Diagnostics | Where-Object Status -eq Failed).Count
    $collectorWarnings=@($script:Diagnostics | Where-Object Status -eq Warning).Count
    $discoveryStatus=if ($collectorFailures) { 'Completed with collector failures' } elseif ($collectorWarnings) { 'Completed with warnings' } else { 'Complete' }

    $summary=@(
        [pscustomobject][ordered]@{Category='Discovery';Metric='Status';Value=$discoveryStatus}
        [pscustomobject][ordered]@{Category='Discovery';Metric='Generated';Value=(Get-Date)}
        [pscustomobject][ordered]@{Category='Discovery';Metric='Collector failures';Value=$collectorFailures}
        [pscustomobject][ordered]@{Category='Discovery';Metric='Collector warnings';Value=$collectorWarnings}
        [pscustomobject][ordered]@{Category='Identity';Metric='Computer name';Value=$ComputerName}
        [pscustomobject][ordered]@{Category='Identity';Metric='Domain';Value=$sys.Domain}
        [pscustomobject][ordered]@{Category='Identity';Metric='Physical or virtual';Value=$sys.PhysicalOrVirtual}
        [pscustomobject][ordered]@{Category='Identity';Metric='Manufacturer';Value=$sys.Manufacturer}
        [pscustomobject][ordered]@{Category='Identity';Metric='Model';Value=$sys.Model}
        [pscustomobject][ordered]@{Category='Identity';Metric='Serial number';Value=$sys.SerialNumber}
        [pscustomobject][ordered]@{Category='Operating system';Metric='Name';Value=$sys.OperatingSystem}
        [pscustomobject][ordered]@{Category='Operating system';Metric='Version';Value=$sys.Version}
        [pscustomobject][ordered]@{Category='Operating system';Metric='Build';Value=$sys.Build}
        [pscustomobject][ordered]@{Category='Operating system';Metric='Last boot';Value=$sys.LastBoot}
        [pscustomobject][ordered]@{Category='Compute';Metric='CPU';Value=$sys.CPU}
        [pscustomobject][ordered]@{Category='Compute';Metric='Sockets';Value=$sys.Sockets}
        [pscustomobject][ordered]@{Category='Compute';Metric='Cores';Value=$sys.Cores}
        [pscustomobject][ordered]@{Category='Compute';Metric='Logical processors';Value=$sys.LogicalProcessors}
        [pscustomobject][ordered]@{Category='Compute';Metric='Memory (GB)';Value=$sys.MemoryGB}
        [pscustomobject][ordered]@{Category='Network';Metric='IPv4 addresses';Value=$ipv4Addresses}
        [pscustomobject][ordered]@{Category='Network';Metric='Gateways';Value=$gateways}
        [pscustomobject][ordered]@{Category='Network';Metric='DNS servers';Value=$dnsServers}
        [pscustomobject][ordered]@{Category='Network';Metric='DHCP servers';Value=(Join-Unique $dhcpClient)}
        [pscustomobject][ordered]@{Category='Network';Metric='Mapped drives';Value=(Join-Unique $mapped)}
        [pscustomobject][ordered]@{Category='Storage';Metric='Volumes';Value=$disks.Count}
        [pscustomobject][ordered]@{Category='Storage';Metric='Total capacity (GB)';Value=$diskTotal}
        [pscustomobject][ordered]@{Category='Storage';Metric='Free capacity (GB)';Value=$diskFree}
        [pscustomobject][ordered]@{Category='Inventory';Metric='Installed roles';Value=$installedRoleCount}
        [pscustomobject][ordered]@{Category='Inventory';Metric='Installed features';Value=$installedFeatureCount}
        [pscustomobject][ordered]@{Category='Inventory';Metric='Applications';Value=$applications.Count}
        [pscustomobject][ordered]@{Category='Inventory';Metric='Services';Value=$services.Count}
        [pscustomobject][ordered]@{Category='Inventory';Metric='Listening ports';Value=$ports.Count}
        [pscustomobject][ordered]@{Category='Inventory';Metric='SMB shares';Value=$shares.Count}
        [pscustomobject][ordered]@{Category='Inventory';Metric='Printers';Value=$printers.Count}
        [pscustomobject][ordered]@{Category='Inventory';Metric='Scheduled tasks';Value=$tasks.Count}
        [pscustomobject][ordered]@{Category='Directory services';Metric='AD users';Value=$adUsers.Count}
        [pscustomobject][ordered]@{Category='Directory services';Metric='AD group memberships';Value=$adMembership.Count}
        [pscustomobject][ordered]@{Category='Directory services';Metric='Replication partners';Value=$adReplication.Count}
        [pscustomobject][ordered]@{Category='Infrastructure services';Metric='DHCP scopes';Value=$dhcp.Count}
        [pscustomobject][ordered]@{Category='Infrastructure services';Metric='DHCP exclusions';Value=$dhcpExclusions.Count}
        [pscustomobject][ordered]@{Category='Infrastructure services';Metric='DHCP reservations';Value=$dhcpReservations.Count}
        [pscustomobject][ordered]@{Category='Infrastructure services';Metric='DHCP options';Value=$dhcpOptions.Count}
        [pscustomobject][ordered]@{Category='Infrastructure services';Metric='DHCP failover relationships';Value=$dhcpFailover.Count}
        [pscustomobject][ordered]@{Category='Infrastructure services';Metric='DNS zones';Value=$dns.Count}
        [pscustomobject][ordered]@{Category='Migration review';Metric='SMB 1.0/CIFS';Value=$(if ($features.Name -contains 'FS-SMB1') { 'Installed' } else { 'Not installed' })}
        [pscustomobject][ordered]@{Category='Migration review';Metric='Service accounts in Windows services';Value=$serviceAccounts.Count}
    )
    $null=Add-DataSheet $book 'Discovery Summary' $summary

    $software=Join-Unique @($applications.Name+$services.Name+$services.DisplayName) ([Environment]::NewLine)
    $azurePatterns=[ordered]@{
        'Microsoft Monitoring Agent'='Microsoft Monitoring Agent|OMS Agent'
        'Azure Data Studio'='Azure Data Studio'
        'Azure Workload Backup'='Azure.*Workload|Microsoft Azure Recovery Services'
        'Azure PowerShell'='Azure PowerShell|Az PowerShell'
        'Power Automate UI flows'='Power Automate|UI flows'
        'On-premises Data Gateway'='On-premises data gateway'
        'Azure plug-in for Veeam'='Veeam.*Azure|Azure.*Veeam'
        'Azure Backup'='Azure Backup'
        'Azure Site Recovery Mobility Service'='Site Recovery.*Mobility|Microsoft Azure Site Recovery'
        'Windows Azure VM Agent'='Windows Azure VM Agent|WindowsAzureGuestAgent|RdAgent'
        'Entra Connect Authentication Agent'='Azure AD Connect Authentication Agent|Entra Connect Authentication Agent'
        'Microsoft Entra Connect'='Azure AD Connect|Microsoft Entra Connect'
        'Citrix Azure Provisioning'='Citrix.*Azure'
        'Azure Connected Storage Service'='Azure Connected Storage'
        'Azure Recovery Services'='Azure Recovery Services'
        'Azure Information Protection'='Azure Information Protection'
        'Defender for Identity Sensor'='Azure ATP|Defender for Identity'
        'Entra Connect Agent Updater'='Azure AD Connect Agent Updater|Entra Connect Agent Updater'
        'NPS Extension for Azure MFA'='NPS Extension.*Azure MFA|Azure MFA.*NPS'
        'Intune Connector for Active Directory'='Intune Connector.*Active Directory'
        'Microsoft Entra Application Proxy Connector'='Application Proxy Connector'
        'Storage Sync Agent'='Storage Sync Agent|Azure File Sync'
    }
    $azureSignals=@(
        [pscustomobject][ordered]@{
            Component='Azure virtual machine'
            Detected=$(if ($sys.Manufacturer -match 'Microsoft' -and $sys.Model -match 'Virtual') { 'Yes' } else { 'No' })
        }
        foreach ($entry in $azurePatterns.GetEnumerator()) {
            [pscustomobject][ordered]@{
                Component=$entry.Key
                Detected=$(if ($software -match $entry.Value) { 'Yes' } else { 'No' })
            }
        }
    )

    $null=Add-DataSheet $book 'System' $system
    $null=Add-DataSheet $book 'Network' $network
    $null=Add-DataSheet $book 'Storage' $disks
    $null=Add-DataSheet $book 'Roles and Features' $features
    $null=Add-DataSheet $book 'Applications' $applications
    $null=Add-DataSheet $book 'Services' $services
    $null=Add-DataSheet $book 'Listening Ports' $ports
    $null=Add-DataSheet $book 'Shares and Permissions' $shares
    $null=Add-DataSheet $book 'Printers' $printers
    $null=Add-DataSheet $book 'Scheduled Tasks' $tasks
    $null=Add-DataSheet $book 'Local Accounts' $accounts
    $null=Add-DataSheet $book 'Local Group Membership' $groups
    $null=Add-DataSheet $book 'Service Accounts' $serviceAccounts
    $null=Add-DataSheet $book 'Azure Components' $azureSignals
    $null=Add-DataSheet $book 'Active Directory' $dc
    $null=Add-DataSheet $book 'AD Users' $adUsers
    $null=Add-DataSheet $book 'AD Group Membership' $adMembership
    $null=Add-DataSheet $book 'AD Replication' $adReplication
    $null=Add-DataSheet $book 'DHCP Scopes' $dhcp
    $null=Add-DataSheet $book 'DHCP Exclusions' $dhcpExclusions
    $null=Add-DataSheet $book 'DHCP Reservations' $dhcpReservations
    $null=Add-DataSheet $book 'DHCP Options' $dhcpOptions
    $null=Add-DataSheet $book 'DHCP Failover' $dhcpFailover
    $null=Add-DataSheet $book 'DNS Zones' $dns
    $null=Add-DataSheet $book 'Diagnostics' $script:Diagnostics

    $book.Worksheets.Delete($workingSheetName)
    Close-ExcelPackage -ExcelPackage $package
    $package=$null
} catch {
    if ($package) { $package.Dispose() }
    throw
}

Write-Host ''
Write-Host 'Server discovery completed.' -ForegroundColor Green
Write-Host ('Output: {0}' -f $outputPath)
Write-Host ('Collector failures: {0}' -f @($script:Diagnostics|Where-Object Status -eq Failed).Count)
Write-Output $outputPath
