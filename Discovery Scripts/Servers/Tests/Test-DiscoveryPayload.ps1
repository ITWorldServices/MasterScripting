#Requires -Version 5.1
<#
.SYNOPSIS
Tests the actual collector payload expressions with synthetic inventory.
.DESCRIPTION
No Windows discovery cmdlets, remoting, AD, WMI, or ImportExcel are used.
Runs on Windows PowerShell 5.1 and PowerShell 7. Extracting the real diagnostics
initializers and return expressions ensures the pre-fix code fails this test.
#>
[CmdletBinding()]
param([string]$SourceDirectory=(Split-Path $PSScriptRoot -Parent))

$ErrorActionPreference='Stop'

function Read-ScriptAst {
    param([string]$Name)
    $tokens=$null
    $parseErrors=$null
    $path=Join-Path $SourceDirectory $Name
    $ast=[System.Management.Automation.Language.Parser]::ParseFile($path,[ref]$tokens,[ref]$parseErrors)
    if($parseErrors.Count){throw ($parseErrors | Out-String)}
    return $ast
}

function Assert-Equal {
    param($Actual,$Expected,[string]$Message)
    if($Actual -ne $Expected){throw ('{0}: expected {1}, got {2}' -f $Message,$Expected,$Actual)}
}

function Test-CollectorPayload {
    param($Body,[string]$Server,[int]$DiagnosticCount)
    $ComputerName=$Server
    $serverClass='General Server'
    $collectionMode='Full role-aware'
    $isSqlServer=$false
    $isExchangeServer=$false
    $hasActiveDirectory=$false
    $featureNames='DNS'

    $initializer=$Body.Find({param($node)
        $node -is [System.Management.Automation.Language.AssignmentStatementAst] -and
        $node.Left.Extent.Text -match '^\$(script:)?Diagnostics$'
    },$true)
    if(-not $initializer){throw 'Diagnostics initializer was not found.'}
    . ([scriptblock]::Create($initializer.Extent.Text))
    # Modern uses script scope; WMI fallback uses function-local scope.
    if($initializer.Left.VariablePath.IsScript){$list=$script:Diagnostics}else{$list=$diagnostics}
    for($i=0;$i -lt $DiagnosticCount;$i++){
        $list.Add([pscustomobject]@{
            ComputerName=$Server; Collector='Synthetic'; Status='Success'
            RecordCount=2; Message=('Diagnostic {0}' -f $i)
        })
    }

    $returnStatement=$Body.EndBlock.Statements[-1]
    $payloadTable=$returnStatement.Find({param($node)
        $node -is [System.Management.Automation.Language.HashtableAst] -and
        @($node.KeyValuePairs | Where-Object {$_.Item1.Value -eq 'Metadata'}).Count -eq 1
    },$true)
    if(-not $payloadTable){throw 'Inventory payload expression was not found.'}
    $sections=@()
    foreach($pair in $payloadTable.KeyValuePairs){
        $key=$pair.Item1.Value
        if($key -in 'Metadata','Diagnostics'){continue}
        $variable=$pair.Item2.Find({param($node)
            $node -is [System.Management.Automation.Language.VariableExpressionAst]
        },$true)
        # The WMI collector intentionally leaves role-specific sections empty.
        if(-not $variable){continue}
        $rows=@()
        if($key -ne 'LocalAccounts'){
            $rows=@(
                [pscustomobject]@{ComputerName=$Server; Name=($key+'-1')}
                [pscustomobject]@{ComputerName=$Server; Name=($key+'-2')}
            )
        }
        Set-Variable -Name $variable.VariablePath.UserPath -Value $rows
        $sections+=@([pscustomobject]@{Name=$key; Count=$rows.Count})
    }

    $output=@(& ([scriptblock]::Create($returnStatement.Extent.Text)))
    Assert-Equal $output.Count 1 'Exactly one inventory payload'
    $payload=$output[0]
    Assert-Equal $payload.Metadata.ComputerName $Server 'Metadata server'
    Assert-Equal ($payload.Diagnostics -is [object[]]) $true 'Diagnostics is a plain array'
    Assert-Equal $payload.Diagnostics.Count $DiagnosticCount 'Diagnostics count'
    foreach($section in $sections){
        $records=$payload.PSObject.Properties[$section.Name].Value
        Assert-Equal ($records -is [object[]]) $true ($section.Name+' is an array')
        Assert-Equal $records.Count $section.Count ($section.Name+' record count')
        foreach($record in $records){Assert-Equal $record.ComputerName $Server 'Record server identity'}
    }
    # Exercise the serialization used to transport remote and job output.
    $xml=[System.Management.Automation.PSSerializer]::Serialize($payload,10)
    $restored=[System.Management.Automation.PSSerializer]::Deserialize($xml)
    Assert-Equal $restored.Diagnostics.Count $DiagnosticCount 'Serialized diagnostics count'
    Assert-Equal $restored.System.Count 2 'Serialized inventory count'
    Assert-Equal $restored.Metadata.ComputerName $Server 'Serialized server identity'
    return $restored
}

$modern=Read-ScriptAst 'Get-ServerDiscoveryData.ps1'
$legacy=Read-ScriptAst 'Get-LegacyServerDiscoveryData.ps1'
$orchestrator=Read-ScriptAst 'Invoke-ADServerDiscovery.ps1'
$legacyFunction=$legacy.Find({param($node)
    $node -is [System.Management.Automation.Language.FunctionDefinitionAst] -and
    $node.Name -eq 'Get-LegacyServerDiscoveryData'
},$true)

$payloads=@()
foreach($kind in 'Modern','Fallback'){
    $body=if($kind -eq 'Modern'){$modern}else{$legacyFunction.Body}
    foreach($count in 0,1,3){
        $payloads+=@(Test-CollectorPayload $body ('{0}-{1}' -f $kind,$count) $count)
    }
}

$rowFunction=$orchestrator.Find({param($node)
    $node -is [System.Management.Automation.Language.FunctionDefinitionAst] -and
    $node.Name -eq 'Get-PayloadRows'
},$true)
. ([scriptblock]::Create($rowFunction.Extent.Text))
foreach($section in 'System','Network','Storage','Applications','Services'){
    $rows=@(Get-PayloadRows $payloads $section)
    Assert-Equal $rows.Count 12 ($section+' consolidated row count')
    Assert-Equal @($rows.ComputerName | Sort-Object -Unique).Count 6 ($section+' consolidated server count')
}
Assert-Equal @(Get-PayloadRows $payloads Diagnostics).Count 8 'Consolidated diagnostics'
Assert-Equal @(Get-PayloadRows $payloads LocalAccounts).Count 0 'Empty sections remain empty'

# Primary failures must retain their original details even after fallback.
foreach($functionName in 'Join-UniqueValue','Add-OrchestratorDiagnostic'){
    $definition=$orchestrator.Find({param($node)
        $node -is [System.Management.Automation.Language.FunctionDefinitionAst] -and
        $node.Name -eq $functionName
    },$true)
    . ([scriptblock]::Create($definition.Extent.Text))
}
$script:Diagnostics=New-Object 'System.Collections.Generic.List[object]'
try{throw 'Synthetic primary collector failure'}catch{$testError=$_}
$testError | Add-Member NoteProperty OriginInfo ([pscustomobject]@{PSComputerName='TEST01.example.test'}) -Force
$remoteErrors=@($testError)
$errorByTarget=@{}
$errorLoop=$orchestrator.Find({param($node)
    $node -is [System.Management.Automation.Language.ForEachStatementAst] -and
    $node.Variable.VariablePath.UserPath -eq 'remoteError'
},$true)
$warnings=@(& { . ([scriptblock]::Create($errorLoop.Extent.Text)) } 3>&1)
Assert-Equal ($errorByTarget['test01.example.test'] -like '*Synthetic primary collector failure*') $true 'FQDN error lookup'
Assert-Equal ($errorByTarget['test01'] -like '*Synthetic primary collector failure*') $true 'Short-name error lookup'
Assert-Equal $script:Diagnostics.Count 1 'Original primary error retained in Diagnostics'
Assert-Equal $warnings.Count 1 'Original primary error printed to console'
Write-Host 'PASS: both collectors, diagnostics counts 0/1/3, serialization, six-server consolidation, and primary error reporting.'
