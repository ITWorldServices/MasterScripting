function Test-StandaloneModernCollector {
    param([version]$PowerShellVersion,[version]$WindowsVersion,[bool]$ForceLegacy=$false)
    return (-not $ForceLegacy -and $PowerShellVersion.Major -ge 4 -and $WindowsVersion -ge [version]'6.2')
}

function Invoke-StandaloneDiscovery {
    param([string]$OutputDirectory='C:\Temp',[bool]$IncludeMicrosoftTasks=$false,[bool]$UseLegacyCollector=$false)
    $ErrorActionPreference='Stop'
    if($env:OS -ne 'Windows_NT'){throw 'Run this script directly on the Windows Server being inventoried.'}
    $identity=[Security.Principal.WindowsIdentity]::GetCurrent()
    $principal=New-Object Security.Principal.WindowsPrincipal -ArgumentList $identity
    if(-not $principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)){
        throw 'Open Windows PowerShell using Run as administrator under an administrator account for this server, then run this script again.'
    }
    if(-not(Test-Path -LiteralPath $OutputDirectory)){New-Item -ItemType Directory -Path $OutputDirectory -Force | Out-Null}
    $OutputDirectory=(Resolve-Path -LiteralPath $OutputDirectory).ProviderPath
    $version=$PSVersionTable.PSVersion
    $osVersion=[Environment]::OSVersion.Version
    $useModern=Test-StandaloneModernCollector $version $osVersion $UseLegacyCollector
    Write-Host ('Discovering {0} locally as {1} (PowerShell {2})...' -f $env:COMPUTERNAME,$identity.Name,$version)
    $primaryError=$null
    if($useModern){
        try{
            # The embedded collector is parsed only on a compatible runtime.
            $payload=& ([scriptblock]::Create($script:StandaloneModernSource)) $IncludeMicrosoftTasks
            if(-not $payload -or -not $payload.PSObject.Properties['Metadata']){throw 'Modern collector did not return an inventory payload.'}
        }catch{
            $primaryError=$_.Exception.Message
            Write-Warning ('Modern local collector failed: {0}. Trying local basic inventory.' -f $primaryError)
            $useModern=$false
        }
    }
    if(-not $useModern){$payload=Get-LocalBasicDiscoveryData -IncludeMicrosoftTasks $IncludeMicrosoftTasks}
    if($primaryError){
        $payload.Diagnostics=@($payload.Diagnostics)+@(New-LocalDiscoveryRow @{Time=Get-Date; ComputerName=$env:COMPUTERNAME; Collector='Modern Local Collection'; Status='Warning'; RecordCount=0; Message=$primaryError})
    }
    $payload.Metadata | Add-Member NoteProperty ExecutionMode 'Local' -Force
    $payload.Metadata | Add-Member NoteProperty RunAs $identity.Name -Force
    $payload.Metadata | Add-Member NoteProperty PowerShellVersion ([string]$version) -Force
    if($env:PROCESSOR_ARCHITEW6432){
        $payload.Diagnostics=@($payload.Diagnostics)+@(New-LocalDiscoveryRow @{Time=Get-Date; ComputerName=$env:COMPUTERNAME; Collector='Process Architecture'; Status='Warning'; RecordCount=0; Message='32-bit PowerShell on a 64-bit OS may omit 64-bit software. Rerun from 64-bit Windows PowerShell for a complete application list.'})
    }
    $prefix=('{0}-ServerDiscovery-{1}' -f $env:COMPUTERNAME,(Get-Date -Format 'yyyyMMdd-HHmmss-fff'))
    $snapshotPath=Join-Path $OutputDirectory ($prefix+'.clixml')
    $outputPath=Join-Path $OutputDirectory ($prefix+'.xlsx')
    # Preserve all collected data before rendering Excel (including long cells).
    $payload | Export-Clixml -Path $snapshotPath -Depth 12 -Encoding UTF8 -NoClobber
    try{Export-LocalDiscoveryWorkbook -Payload $payload -Path $outputPath}
    catch{throw ('Excel export failed: {0}. Collected inventory is retained at {1}' -f $_.Exception.Message,$snapshotPath)}
    $failed=@($payload.Diagnostics | Where-Object {$_.Status -eq 'Failed'}).Count
    Write-Host ('Local discovery completed. Collector failures: {0}' -f $failed)
    Write-Host ('Output: {0}' -f $outputPath)
    Write-Host ('Inventory snapshot: {0}' -f $snapshotPath)
    Write-Output $outputPath
}

if($MyInvocation.InvocationName -ne '.'){
    Invoke-StandaloneDiscovery -OutputDirectory $OutputDirectory -IncludeMicrosoftTasks ([bool]$IncludeMicrosoftTasks) -UseLegacyCollector ([bool]$UseLegacyCollector)
}
