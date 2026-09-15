#Requires -Version 5.1
[CmdletBinding()]
param([switch]$Check)
$ErrorActionPreference='Stop'
$serverDirectory=Split-Path $PSScriptRoot -Parent
$modern=[IO.File]::ReadAllText((Join-Path $serverDirectory 'Get-ServerDiscoveryData.ps1'))
$zip=[IO.File]::ReadAllText((Join-Path $PSScriptRoot 'ZipWriter.cs'))
if($modern -match "(?m)^'@" -or $zip -match "(?m)^'@"){throw 'Embedded source contains a here-string terminator.'}
$content=[IO.File]::ReadAllText((Join-Path $PSScriptRoot 'Header.ps1'))
foreach($name in 'LocalCollector.ps1','WorkbookWriter.ps1'){$content+="`n"+[IO.File]::ReadAllText((Join-Path $PSScriptRoot $name))}
$content+="`n`$script:StandaloneModernSource=@'`n"+$modern.TrimEnd()+"`n'@`n"
$content+="`n`$script:StandaloneZipSource=@'`n"+$zip.TrimEnd()+"`n'@`n"
$content+="`n"+[IO.File]::ReadAllText((Join-Path $PSScriptRoot 'Main.ps1'))
$content=$content.Replace("`r`n","`n").TrimEnd()+"`n"
$path=Join-Path $serverDirectory 'Invoke-StandaloneServerDiscovery.ps1'
if($Check){
    if(-not(Test-Path $path) -or [IO.File]::ReadAllText($path).Replace("`r`n","`n") -cne $content){throw 'Standalone script is out of date. Rebuild it.'}
    Write-Host 'PASS: standalone script matches its sources and the shared modern collector.'
}else{
    [IO.File]::WriteAllText($path,$content,(New-Object Text.UTF8Encoding -ArgumentList $true))
    Write-Host ('Built {0}' -f $path)
}
