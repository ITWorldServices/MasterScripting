#Requires -Version 5.1
<#
.SYNOPSIS
Inventories every enabled Windows Server computer in Active Directory.

.DESCRIPTION
Runs Get-ServerDiscoveryData.ps1 remotely with PowerShell remoting and creates
one consolidated Excel workbook locally. SQL Server and Exchange Server hosts
are automatically limited by the collector to basic inventory. AD DS, DHCP,
and DNS discovery runs only on general servers that have the matching role.

.EXAMPLE
.\Invoke-ADServerDiscovery.ps1

.EXAMPLE
.\Invoke-ADServerDiscovery.ps1 -SearchBase 'OU=Servers,DC=contoso,DC=com' -ThrottleLimit 12
#>
[CmdletBinding()]
param(
    [string]$SearchBase,
    [string[]]$IncludeComputerName,
    [string[]]$ExcludeComputerName,
    [ValidateRange(1,64)][int]$ThrottleLimit=8,
    [ValidateRange(5,300)][int]$ConnectionTimeoutSeconds=30,
    [ValidateRange(30,3600)][int]$OperationTimeoutSeconds=600,
    [pscredential]$Credential,
    [switch]$IncludeMicrosoftTasks,
    [string]$OutputDirectory='C:\Temp'
)

$ErrorActionPreference='Stop'
$ProgressPreference='SilentlyContinue'
$script:Diagnostics=[System.Collections.Generic.List[object]]::new()

function Add-OrchestratorDiagnostic {
    param([string]$ComputerName,[string]$Status,[string]$Message)
    $script:Diagnostics.Add([pscustomobject][ordered]@{
        Time=Get-Date
        ComputerName=$ComputerName
        Collector='Remote Collection'
        Status=$Status
        RecordCount=0
        Message=$Message
    })
}

function Join-UniqueValue {
    param([object[]]$Values,[string]$Separator='; ')
    $joined=(($Values | Where-Object {
        $null -ne $_ -and -not [string]::IsNullOrWhiteSpace([string]$_)
    } | ForEach-Object {([string]$_).Trim()} | Sort-Object -Unique) -join $Separator)
    if([string]::IsNullOrWhiteSpace($joined)){return $null}
    return $joined
}

function Convert-ToCellValue {
    param($Value)
    if($null -eq $Value){return $null}
    if($Value -is [bool]){return $(if($Value){'Yes'}else{'No'})}
    if($Value -is [array]){return (Join-UniqueValue $Value)}
    if($Value -is [datetime] -or $Value -is [string] -or $Value -is [decimal] -or $Value.GetType().IsPrimitive){return $Value}
    return [string]$Value
}

function Test-ComputerPattern {
    param($Computer,[string[]]$Patterns)
    foreach($pattern in @($Patterns)){
        if($Computer.Name -like $pattern -or $Computer.DNSHostName -like $pattern){return $true}
    }
    return $false
}

function Get-PayloadRows {
    param([object[]]$Payloads,[string]$PropertyName)
    foreach($payload in @($Payloads)){
        foreach($row in @($payload.PSObject.Properties[$PropertyName].Value)){
            if($null -ne $row){$row}
        }
    }
}

function Initialize-ImportExcel {
    $requiredVersion='7.8.10'
    $available=Get-Module -ListAvailable ImportExcel | Where-Object {$_.Version -ge [version]$requiredVersion} | Select-Object -First 1
    if($available){Import-Module ImportExcel -MinimumVersion $requiredVersion -ErrorAction Stop; return}

    Write-Host ('ImportExcel {0} is not installed. Bootstrapping PowerShell Gallery access...' -f $requiredVersion) -ForegroundColor Yellow
    try{
        [Net.ServicePointManager]::SecurityProtocol=[Net.ServicePointManager]::SecurityProtocol -bor [Net.SecurityProtocolType]::Tls12
        $nuget=Get-PackageProvider NuGet -ListAvailable -ErrorAction SilentlyContinue | Where-Object {$_.Version -ge [version]'2.8.5.201'} | Select-Object -First 1
        if(-not $nuget){Install-PackageProvider NuGet -MinimumVersion 2.8.5.201 -Force -Confirm:$false | Out-Null}
        $gallery=Get-PSRepository PSGallery -ErrorAction SilentlyContinue
        if(-not $gallery){Register-PSRepository -Default -ErrorAction Stop; $gallery=Get-PSRepository PSGallery}
        $originalPolicy=$gallery.InstallationPolicy
        if($originalPolicy -ne 'Trusted'){Set-PSRepository PSGallery -InstallationPolicy Trusted}
        try{
            Install-Module ImportExcel -RequiredVersion $requiredVersion -Repository PSGallery -Scope CurrentUser -Force -AllowClobber -Confirm:$false -ErrorAction Stop
        }finally{
            if($originalPolicy -and $originalPolicy -ne 'Trusted'){Set-PSRepository PSGallery -InstallationPolicy $originalPolicy -ErrorAction SilentlyContinue}
        }
        Import-Module ImportExcel -RequiredVersion $requiredVersion -Force -ErrorAction Stop
    }catch{
        throw ('Unable to install ImportExcel. Confirm HTTPS access to PowerShell Gallery and NuGet, including proxy configuration. Error: {0}' -f $_.Exception.Message)
    }
}

function Set-WorksheetStyle {
    param($Worksheet)
    if(-not $Worksheet -or -not $Worksheet.Dimension){return}
    $lastRow=$Worksheet.Dimension.End.Row
    $lastColumn=$Worksheet.Dimension.End.Column
    $header=$Worksheet.Cells[1,1,1,$lastColumn]
    $header.Style.Font.Bold=$true
    $header.Style.Font.Color.SetColor([System.Drawing.Color]::White)
    $header.Style.Fill.PatternType=[OfficeOpenXml.Style.ExcelFillStyle]::Solid
    $header.Style.Fill.BackgroundColor.SetColor([System.Drawing.Color]::FromArgb(31,78,121))
    $header.Style.HorizontalAlignment=[OfficeOpenXml.Style.ExcelHorizontalAlignment]::Center
    $header.Style.VerticalAlignment=[OfficeOpenXml.Style.ExcelVerticalAlignment]::Center
    $header.Style.WrapText=$true
    $header.AutoFilter=$true
    $Worksheet.Cells[$Worksheet.Dimension.Address].Style.Font.Name='Arial'
    $Worksheet.Cells[$Worksheet.Dimension.Address].Style.Font.Size=10
    $Worksheet.Cells[$Worksheet.Dimension.Address].AutoFitColumns()
    for($column=1;$column -le $lastColumn;$column++){
        if($Worksheet.Column($column).Width -lt 12){$Worksheet.Column($column).Width=12}
        if($Worksheet.Column($column).Width -gt 45){
            $Worksheet.Column($column).Width=45
            if($lastRow -gt 1){$Worksheet.Cells[2,$column,$lastRow,$column].Style.WrapText=$true}
        }
    }
    $Worksheet.View.ShowGridLines=$false
    $Worksheet.View.FreezePanes(2,1)
}

function Add-DataSheet {
    param($Workbook,[string]$Name,[object[]]$Rows)
    if($Workbook.Worksheets[$Name]){$Workbook.Worksheets.Delete($Name)}
    $sheet=$Workbook.Worksheets.Add($Name)
    $data=@($Rows)
    if(-not $data.Count){
        $sheet.Cells[1,1].Value='Status'
        $sheet.Cells[2,1].Value='No records returned.'
        Set-WorksheetStyle $sheet
        return
    }
    $headers=@($data[0].PSObject.Properties.Name)
    for($column=0;$column -lt $headers.Count;$column++){$sheet.Cells[1,($column+1)].Value=$headers[$column]}
    for($row=0;$row -lt $data.Count;$row++){
        for($column=0;$column -lt $headers.Count;$column++){
            $propertyName=$headers[$column]
            $property=$data[$row].PSObject.Properties[$propertyName]
            $value=Convert-ToCellValue $(if($property){$property.Value}else{$null})
            $cell=$sheet.Cells[($row+2),($column+1)]
            $cell.Value=$value
            if($value -is [datetime]){$cell.Style.Numberformat.Format='yyyy-mm-dd HH:mm:ss'}
            elseif($propertyName -match 'Percent|Utilization|SizeGB|FreeGB|MemoryGB'){$cell.Style.Numberformat.Format='0.00'}
        }
    }
    Set-WorksheetStyle $sheet
}

if($env:OS -ne 'Windows_NT'){throw 'Run this script from a domain-connected Windows management server.'}
if(-not(Get-Module -ListAvailable ActiveDirectory)){throw 'The ActiveDirectory PowerShell module is required on the initiating server.'}
Import-Module ActiveDirectory -ErrorAction Stop
Initialize-ImportExcel

$collectorPath=Join-Path $PSScriptRoot 'Get-ServerDiscoveryData.ps1'
if(-not(Test-Path -LiteralPath $collectorPath)){throw ('Remote collector was not found: {0}' -f $collectorPath)}
if(-not(Test-Path -LiteralPath $OutputDirectory)){New-Item -ItemType Directory -Path $OutputDirectory -Force | Out-Null}

$adParameters=@{
    Filter={Enabled -eq $true -and OperatingSystem -like '*Server*'}
    Properties='DNSHostName','OperatingSystem','LastLogonDate','DistinguishedName'
}
if($SearchBase){$adParameters.SearchBase=$SearchBase}
$adComputers=@(Get-ADComputer @adParameters | Sort-Object Name)
if($IncludeComputerName){$adComputers=@($adComputers | Where-Object {Test-ComputerPattern $_ $IncludeComputerName})}
if($ExcludeComputerName){$adComputers=@($adComputers | Where-Object {-not(Test-ComputerPattern $_ $ExcludeComputerName)})}
if(-not $adComputers.Count){throw 'No enabled Windows Server computer objects matched the requested AD scope and filters.'}

$targets=@($adComputers | ForEach-Object {if($_.DNSHostName){$_.DNSHostName}else{$_.Name}})
Write-Host ('Discovered {0} enabled Windows Server computer objects in AD.' -f $targets.Count) -ForegroundColor Cyan
Write-Host ('Starting remote collection with throttle limit {0}...' -f $ThrottleLimit) -ForegroundColor Cyan

$sessionOption=New-PSSessionOption -OpenTimeout ($ConnectionTimeoutSeconds*1000) -OperationTimeout ($OperationTimeoutSeconds*1000)
$invokeParameters=@{
    ComputerName=$targets
    FilePath=$collectorPath
    ArgumentList=[bool]$IncludeMicrosoftTasks.IsPresent
    ThrottleLimit=$ThrottleLimit
    SessionOption=$sessionOption
    ErrorAction='SilentlyContinue'
    ErrorVariable='remoteErrors'
}
if($Credential){$invokeParameters.Credential=$Credential}
$payloads=@(Invoke-Command @invokeParameters | Where-Object {$_.PSObject.Properties['Metadata'] -and $_.PSObject.Properties['System']})

$payloadByTarget=@{}
foreach($payload in $payloads){
    $key=[string]$payload.PSComputerName
    if([string]::IsNullOrWhiteSpace($key)){$key=[string]$payload.Metadata.ComputerName}
    if($key){$payloadByTarget[$key.ToLowerInvariant()]=$payload}
}
$errorByTarget=@{}
foreach($remoteError in @($remoteErrors)){
    $errorTarget=[string]$remoteError.OriginInfo.PSComputerName
    if([string]::IsNullOrWhiteSpace($errorTarget)){$errorTarget='Unknown'}
    $errorByTarget[$errorTarget.ToLowerInvariant()]=[string]$remoteError.Exception.Message
}

$serverSummary=@()
foreach($computer in $adComputers){
    $target=if($computer.DNSHostName){$computer.DNSHostName}else{$computer.Name}
    $key=$target.ToLowerInvariant()
    $payload=$payloadByTarget[$key]
    if(-not $payload){
        $payload=$payloads | Where-Object {$_.Metadata.ComputerName -eq $computer.Name} | Select-Object -First 1
    }
    if($payload){
        $failed=@($payload.Diagnostics | Where-Object Status -eq Failed)
        $inventoryStatus=if($failed.Count){'Completed with collector failures'}else{'Complete'}
        Add-OrchestratorDiagnostic $computer.Name Success ('Remote collection completed as {0}.' -f $payload.Metadata.CollectionMode)
        $errorMessage=Join-UniqueValue $failed.Message
        $meta=$payload.Metadata
    }else{
        $inventoryStatus='Unavailable'
        $errorMessage=$errorByTarget[$key]
        if([string]::IsNullOrWhiteSpace($errorMessage)){$errorMessage='No inventory payload was returned. Verify DNS, WinRM, firewall, and administrative access.'}
        Add-OrchestratorDiagnostic $computer.Name Failed $errorMessage
        $meta=$null
    }
    $serverSummary+=[pscustomobject][ordered]@{
        ComputerName=$computer.Name
        DNSHostName=$computer.DNSHostName
        ADOperatingSystem=$computer.OperatingSystem
        DistinguishedName=$computer.DistinguishedName
        ADLastLogonDate=$computer.LastLogonDate
        InventoryStatus=$inventoryStatus
        ServerClass=$(if($meta){$meta.ServerClass}else{'Unknown'})
        CollectionMode=$(if($meta){$meta.CollectionMode}else{'Not collected'})
        HasActiveDirectory=$(if($meta){$meta.HasActiveDirectory}else{$null})
        HasDhcp=$(if($meta){$meta.HasDhcp}else{$null})
        HasDns=$(if($meta){$meta.HasDns}else{$null})
        Error=$errorMessage
    }
}

$adUsers=@(Get-PayloadRows $payloads ADUsers | Sort-Object Domain,DistinguishedName -Unique)
$adMembership=@(Get-PayloadRows $payloads ADGroupMembership | Sort-Object Domain,Username,GroupName,MembershipType -Unique)
$diagnostics=@($script:Diagnostics)+@(Get-PayloadRows $payloads Diagnostics)

$domain=try{Get-ADDomain}catch{$null}
$prefix=if($domain){$domain.NetBIOSName}else{'AD'}
$safePrefix=$prefix -replace '[^A-Za-z0-9_-]','_'
$outputPath=Join-Path $OutputDirectory ('{0}-ServerDiscovery-{1}.xlsx' -f $safePrefix,(Get-Date -Format yyyyMMdd-HHmmss))

$package=$null
try{
    $package=Open-ExcelPackage -Path $outputPath -Create
    $book=$package.Workbook
    Add-DataSheet $book 'Server Summary' $serverSummary
    Add-DataSheet $book 'System' @(Get-PayloadRows $payloads System)
    Add-DataSheet $book 'Network' @(Get-PayloadRows $payloads Network)
    Add-DataSheet $book 'Storage' @(Get-PayloadRows $payloads Storage)
    Add-DataSheet $book 'Roles and Features' @(Get-PayloadRows $payloads RolesAndFeatures)
    Add-DataSheet $book 'Applications' @(Get-PayloadRows $payloads Applications)
    Add-DataSheet $book 'Services' @(Get-PayloadRows $payloads Services)
    Add-DataSheet $book 'Listening Ports' @(Get-PayloadRows $payloads ListeningPorts)
    Add-DataSheet $book 'Shares and Permissions' @(Get-PayloadRows $payloads SharesAndPermissions)
    Add-DataSheet $book 'Scheduled Tasks' @(Get-PayloadRows $payloads ScheduledTasks)
    Add-DataSheet $book 'Local Accounts' @(Get-PayloadRows $payloads LocalAccounts)
    Add-DataSheet $book 'Local Group Membership' @(Get-PayloadRows $payloads LocalGroupMembership)
    Add-DataSheet $book 'Service Accounts' @(Get-PayloadRows $payloads ServiceAccounts)
    Add-DataSheet $book 'Active Directory' @(Get-PayloadRows $payloads ActiveDirectory)
    Add-DataSheet $book 'AD Users' $adUsers
    Add-DataSheet $book 'AD Group Membership' $adMembership
    Add-DataSheet $book 'AD Replication' @(Get-PayloadRows $payloads ADReplication)
    Add-DataSheet $book 'DHCP Scopes' @(Get-PayloadRows $payloads DHCPScopes)
    Add-DataSheet $book 'DHCP Exclusions' @(Get-PayloadRows $payloads DHCPExclusions)
    Add-DataSheet $book 'DHCP Reservations' @(Get-PayloadRows $payloads DHCPReservations)
    Add-DataSheet $book 'DHCP Options' @(Get-PayloadRows $payloads DHCPOptions)
    Add-DataSheet $book 'DHCP Failover' @(Get-PayloadRows $payloads DHCPFailover)
    Add-DataSheet $book 'DNS Zones' @(Get-PayloadRows $payloads DNSZones)
    Add-DataSheet $book 'Diagnostics' $diagnostics
    Close-ExcelPackage -ExcelPackage $package
    $package=$null
}catch{
    if($package){$package.Dispose()}
    throw
}

$unavailable=@($serverSummary | Where-Object InventoryStatus -eq Unavailable).Count
$collectorFailures=@($diagnostics | Where-Object Status -eq Failed).Count
Write-Host ''
Write-Host 'AD-wide server discovery completed.' -ForegroundColor Green
Write-Host ('Servers targeted: {0}' -f $serverSummary.Count)
Write-Host ('Servers unavailable: {0}' -f $unavailable)
Write-Host ('Collector failures: {0}' -f $collectorFailures)
Write-Host ('Output: {0}' -f $outputPath)
Write-Output $outputPath
