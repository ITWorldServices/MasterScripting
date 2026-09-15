#Requires -Version 5.1
[CmdletBinding()]
param([string]$OutputDirectory=(Join-Path ([IO.Path]::GetTempPath()) ('standalone-discovery-test-'+[guid]::NewGuid())))
$ErrorActionPreference='Stop'
$serverDirectory=Split-Path $PSScriptRoot -Parent
& (Join-Path $serverDirectory 'Standalone/Build-StandaloneServerDiscovery.ps1') -Check
$path=Join-Path $serverDirectory 'Invoke-StandaloneServerDiscovery.ps1'
$tokens=$null; $errors=$null
$ast=[Management.Automation.Language.Parser]::ParseFile($path,[ref]$tokens,[ref]$errors)
if($errors.Count){throw ($errors | Out-String)}
$testOutputDirectory=$OutputDirectory
. $path
$OutputDirectory=$testOutputDirectory

function Assert-Standalone {
    param([bool]$Condition,[string]$Message)
    if(-not $Condition){throw $Message}
}
Assert-Standalone (-not(Test-StandaloneModernCollector '2.0' '6.1')) 'Server 2008 R2 / PS2 must use local basic inventory.'
Assert-Standalone (-not(Test-StandaloneModernCollector '5.1' '6.1')) 'Server 2008 R2 / PS5.1 must use local basic inventory.'
Assert-Standalone (Test-StandaloneModernCollector '4.0' '6.3') 'Server 2012 R2 / PS4 must use modern local inventory.'
Assert-Standalone (Test-StandaloneModernCollector '5.1' '10.0') 'Modern Windows must use modern local inventory.'
Assert-Standalone (-not(Test-StandaloneModernCollector '5.1' '10.0' $true)) 'Forced local basic selection failed.'

# Reject newer PowerShell syntax outside the intentionally embedded modern string.
$incompatible=$ast.FindAll({param($node)
    ($node -is [Management.Automation.Language.TypeExpressionAst] -and $node.TypeName.FullName -match '^(pscustomobject|ordered)$') -or
    ($node -is [Management.Automation.Language.BinaryExpressionAst] -and $node.Operator -match '^(I|C)?(Not)?In$') -or
    ($node -is [Management.Automation.Language.InvokeMemberExpressionAst] -and $node.Static -and $node.Member.Value -eq 'new')
},$true)
Assert-Standalone ($incompatible.Count -eq 0) 'Found newer PowerShell syntax in the legacy/bootstrap path.'
$wmiCalls=$ast.FindAll({param($node) $node -is [Management.Automation.Language.CommandAst] -and $node.GetCommandName() -eq 'Get-WmiObject'},$true)
Assert-Standalone ($wmiCalls.Count -gt 0) 'Local WMI collector is missing.'
foreach($call in $wmiCalls){Assert-Standalone ($call.Extent.Text -notmatch '-ComputerName|-Credential') 'WMI collector must use local calls.'}

$script:LocalDiagnostics=New-Object 'System.Collections.Generic.List[object]'
$partial=Invoke-LocalDiscoverySection 'Partial test' {New-LocalDiscoveryRow @{ComputerName='TEST01'; Name='Retained row'}; throw 'Synthetic failure after first row'} -WarningAction SilentlyContinue
Assert-Standalone ($partial.Count -eq 1) 'A later failure discarded earlier rows.'
Assert-Standalone ($script:LocalDiagnostics[0].Status -eq 'Failed' -and $script:LocalDiagnostics[0].RecordCount -eq 1) 'Partial collector diagnostic is inaccurate.'
$empty=Invoke-LocalDiscoverySection 'Empty test' {} -WarningAction SilentlyContinue
Assert-Standalone ($empty -is [array] -and $empty.Count -eq 0) 'Empty section is not an empty array.'

# Execute the legacy payload builder against synthetic WMI/registry providers.
# Platform-only COM, ACL, WinNT and netstat paths deliberately fail in this test:
# the remaining inventory must survive and the missing sections must be recorded.
function Test-LocalBasicBuilder {
    $previousComputerName=$env:COMPUTERNAME
    try {
    $env:COMPUTERNAME='TEST01'
    $realSection=${function:Invoke-LocalDiscoverySection}
    function Invoke-LocalDiscoverySection {
        param([string]$Name,[scriptblock]$Action)
        if($Name -in 'Listening Ports','Shares and Permissions','Scheduled Tasks','Local Group Membership'){
            $Action={throw 'Synthetic unavailable Windows interface'}
        }
        & $realSection $Name $Action
    }
    function Get-WmiObject {
        [CmdletBinding()]
        param([string]$Class,[string]$Filter)
        switch($Class){
            'Win32_ComputerSystem' {New-LocalDiscoveryRow @{Name=$env:COMPUTERNAME; Domain='WORKGROUP'; DomainRole=3; Manufacturer='Example'; Model='Physical Server'; TotalPhysicalMemory=8GB}}
            'Win32_OperatingSystem' {New-LocalDiscoveryRow @{Caption='Windows Server 2008 R2'; Version='6.1'; BuildNumber='7601'; LastBootUpTime=$null}}
            'Win32_Processor' {New-LocalDiscoveryRow @{Name='CPU'; NumberOfCores=2; NumberOfLogicalProcessors=4}}
            'Win32_BIOS' {New-LocalDiscoveryRow @{SerialNumber='000123'}}
            'Win32_NetworkAdapterConfiguration' {New-LocalDiscoveryRow @{Description='NIC'; IPAddress=@('192.0.2.1'); IPSubnet=@('255.255.255.0'); DefaultIPGateway=@('192.0.2.254'); DNSServerSearchOrder=@('192.0.2.53'); MACAddress='00:11:22:33:44:55'}}
            'Win32_LogicalDisk' {New-LocalDiscoveryRow @{DeviceID='C:'; VolumeName='OS'; FileSystem='NTFS'; Size=100GB; FreeSpace=25GB}}
            'Win32_ServerFeature' {New-LocalDiscoveryRow @{Name='File Services'}}
            'Win32_Service' {New-LocalDiscoveryRow @{Name='MSSQL$APP'; DisplayName='Application SQL'; State='Running'; StartMode='Auto'; StartName='EXAMPLE\svc'; PathName='C:\Example\sqlservr.exe'}}
            'Win32_UserAccount' {New-LocalDiscoveryRow @{Name='LocalAdmin'; Disabled=$false; Description='Example'; PasswordExpires=$false}}
            default {throw ('Unexpected WMI query: '+$Class)}
        }
    }
    function Test-Path {param([string]$Path) return ($Path -like 'HKLM:*')}
    function Get-ChildItem {param([string]$Path) New-LocalDiscoveryRow @{PSPath=($Path+'\App'); PSChildName='App'}}
    function Get-ItemProperty {param([string]$Path) New-LocalDiscoveryRow @{DisplayName='Application'; DisplayVersion='1.0'; SystemComponent=0; Publisher='Example'}}
    $basic=Get-LocalBasicDiscoveryData
    Assert-Standalone ($basic.System.Count -eq 1 -and $basic.System[0].MemoryGB -eq 8) 'Local WMI system data was lost.'
    Assert-Standalone ($basic.Network[0].PrefixLength -eq 24) 'Legacy subnet mask conversion failed.'
    Assert-Standalone ($basic.Storage[0].PercentFree -eq 25) 'Legacy storage calculation failed.'
    Assert-Standalone ($basic.LocalAccounts.Count -eq 1) 'Workgroup local accounts were omitted.'
    Assert-Standalone ($basic.Metadata.ServerClass -eq 'SQL Server') 'Legacy SQL classification failed.'
    Assert-Standalone (@($basic.Diagnostics | Where-Object {$_.Status -eq 'Failed'}).Count -eq 4) 'Unavailable native interfaces were not diagnosed.'
    Assert-Standalone ($basic.ADUsers.Count -eq 0 -and $basic.DNSZones.Count -eq 0) 'Basic-only host unexpectedly received role-specific rows.'
    Assert-Standalone ($basic.Diagnostics -is [object[]]) 'Local diagnostics array conversion failed.'
    Assert-Standalone ($basic.System[0].ComputerName -eq 'TEST01') 'Local host identity was lost.'
    } finally {$env:COMPUTERNAME=$previousComputerName}
}
Test-LocalBasicBuilder

$payload=New-LocalDiscoveryRow @{
    Metadata=(New-LocalDiscoveryRow @{ComputerName='TEST01'; ServerClass='General Server'; CollectionMode='Local legacy basic only'; HasActiveDirectory=$false; HasDhcp=$false; HasDns=$false; CollectedAt=[datetime]'2026-01-02T03:04:05'})
    System=@(New-LocalDiscoveryRow @{ComputerName='TEST01'; OperatingSystem='Windows Server 2008 R2'; Domain='WORKGROUP'; MemoryGB=16.25; LastBoot=[datetime]'2026-01-01T01:02:03'; SerialNumber='001234'})
    Network=@(New-LocalDiscoveryRow @{ComputerName='TEST01'; IPAddress='192.0.2.10'; MacAddress='00:11:22:33:44:55'})
    Storage=@(New-LocalDiscoveryRow @{ComputerName='TEST01'; Drive='C:'; SizeGB=120.5; FreeGB=30.25; PercentFree=25.1})
    Applications=@(
        (New-LocalDiscoveryRow @{ComputerName='TEST01'; Name='Tool <A> & B'; Version='001.02'; Publisher='Example'})
        (New-LocalDiscoveryRow @{ComputerName='TEST01'; Name='=HYPERLINK("https://example.test")'; ExtraField='Preserve a field absent in the first row'; Notes=('_x0041_' + [char]1)})
    )
    Services=@(New-LocalDiscoveryRow @{ComputerName='TEST01'; Name='Example'; State='Running'; StartName='LocalSystem'})
    LocalAccounts=@(); Diagnostics=$script:LocalDiagnostics.ToArray()
}
if(-not(Test-Path $OutputDirectory)){New-Item -ItemType Directory -Path $OutputDirectory | Out-Null}
$xlsx=Join-Path $OutputDirectory 'Standalone-fixture.xlsx'
$clixml=Join-Path $OutputDirectory 'Standalone-fixture.clixml'
$payload | Export-Clixml -Path $clixml -Depth 12
Export-LocalDiscoveryWorkbook $payload $xlsx
$restored=Import-Clixml $clixml
Assert-Standalone ($restored.Applications.Count -eq 2 -and $restored.System[0].SerialNumber -eq '001234') 'Raw inventory snapshot changed values.'
Add-Type -AssemblyName System.IO.Compression.FileSystem
$archive=[IO.Compression.ZipFile]::OpenRead($xlsx)
try{
    $documents=@{}
    foreach($entry in $archive.Entries){
        $reader=New-Object IO.StreamReader -ArgumentList ($entry.Open())
        try{$documents[$entry.FullName]=[xml]$reader.ReadToEnd()}finally{$reader.Dispose()}
    }
    $sheets=$documents['xl/workbook.xml'].workbook.sheets.sheet
    Assert-Standalone ($sheets.Count -eq 24) 'Expected summary, 22 inventory sections, and Diagnostics.'
    Assert-Standalone (@($sheets | Where-Object {$_.name -eq 'Sheet1'}).Count -eq 0) 'Workbook contains an unwanted blank tab.'
    foreach($item in $sheets){
        $sheet=$documents['xl/worksheets/sheet'+$item.sheetId+'.xml'].worksheet
        Assert-Standalone ($sheet.sheetViews.sheetView.pane.state -eq 'frozen') 'Header row is not frozen.'
        Assert-Standalone ($null -ne $sheet.autoFilter) 'Sheet is missing its filter.'
    }
    $apps=$documents['xl/worksheets/sheet6.xml'].worksheet
    $allCells=@($apps.sheetData.row | ForEach-Object {$_.c})
    Assert-Standalone (@($allCells | Where-Object { $_.is.t.'#text' -eq 'Tool <A> & B' }).Count -eq 1) 'XML text escaping lost data.'
    Assert-Standalone (@($allCells | Where-Object { $_.is.t.'#text' -eq 'ExtraField' }).Count -eq 1) 'Exporter dropped fields absent from the first row.'
    Assert-Standalone (@($allCells | Where-Object {$_.f}).Count -eq 0) 'Inventory text became an Excel formula.'
    $systemCells=@($documents['xl/worksheets/sheet2.xml'].worksheet.sheetData.row | ForEach-Object {$_.c})
    Assert-Standalone (@($systemCells | Where-Object {$_.v -eq '16.25'}).Count -eq 1) 'Numeric cells are not invariant-culture numbers.'
    Assert-Standalone (@($systemCells | Where-Object {$_.is.t.'#text' -eq '001234'}).Count -eq 1) 'Serial number lost its leading zeros.'
}finally{$archive.Dispose()}
Write-Host 'PASS: collector selection, bootstrap syntax, local WMI calls, partial failures, snapshot, and XLSX structure/data.'
Write-Host ('Fixture workbook: {0}' -f $xlsx)
