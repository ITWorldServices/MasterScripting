function Convert-DiscoveryXmlText {
    param($Value)
    $text=[string]$Value
    # Escape SpreadsheetML's literal _xNNNN_ sequences as well as XML markup.
    $text=$text -replace '_x([0-9a-fA-F]{4})_', '_x005F_x$1_'
    $text=$text -replace '[\x00-\x08\x0B\x0C\x0E-\x1F]', ''
    [Security.SecurityElement]::Escape($text)
}

function Get-DiscoveryColumnName {
    param([int]$Number)
    $name=''
    while($Number -gt 0){$Number--; $name=[string][char](65+($Number%26))+$name; $Number=[int][math]::Floor($Number/26)}
    return $name
}

function Convert-DiscoveryCellXml {
    param([string]$Address,$Value,[int]$Style=0)
    if($null -eq $Value){return ('<c r="{0}" s="{1}"/>' -f $Address,$Style)}
    if($Value -is [datetime]){
        return ('<c r="{0}" s="2"><v>{1}</v></c>' -f $Address,$Value.ToOADate().ToString('R',[Globalization.CultureInfo]::InvariantCulture))
    }
    if($Value -is [bool]){$Value=if($Value){'Yes'}else{'No'}}
    elseif($Value -is [array]){$Value=Join-LocalDiscoveryValue $Value}
    elseif($Value -is [ValueType] -and -not($Value -is [enum]) -and -not($Value -is [timespan])){
        $numeric=0.0
        if([double]::TryParse([string]$Value,[ref]$numeric) -and -not [double]::IsNaN($numeric) -and -not [double]::IsInfinity($numeric)){
            return ('<c r="{0}" s="{1}"><v>{2}</v></c>' -f $Address,$Style,([IFormattable]$Value).ToString($null,[Globalization.CultureInfo]::InvariantCulture))
        }
    }
    $text=[string]$Value
    if($text.Length -gt 32767){
        $text=$text.Substring(0,32700)+' [truncated; full value in CLIXML snapshot]'
        Write-Warning ('Excel cell {0} exceeds 32,767 characters; the snapshot retains the full value.' -f $Address)
    }
    return ('<c r="{0}" s="{1}" t="inlineStr"><is><t xml:space="preserve">{2}</t></is></c>' -f $Address,$Style,(Convert-DiscoveryXmlText $text))
}

function Export-LocalDiscoveryWorkbook {
    param($Payload,[string]$Path)
    if(-not ('StandaloneDiscoveryZip' -as [type])){Add-Type -TypeDefinition $script:StandaloneZipSource -ErrorAction Stop}
    $metadata=$Payload.Metadata
    $failed=@($Payload.Diagnostics | Where-Object {$_.Status -eq 'Failed'})
    $warnings=@($Payload.Diagnostics | Where-Object {$_.Status -eq 'Warning'})
    $status=if($failed.Count){'Completed with collector failures'}elseif($warnings.Count){'Completed with warnings'}else{'Complete'}
    $summary=New-LocalDiscoveryRow @{
        ComputerName=$metadata.ComputerName; InventoryStatus=$status; ServerClass=$metadata.ServerClass
        CollectionMode=$metadata.CollectionMode; ExecutionMode='Local'; CollectedAt=$metadata.CollectedAt
        HasActiveDirectory=$metadata.HasActiveDirectory; HasDhcp=$metadata.HasDhcp; HasDns=$metadata.HasDns
        CollectorFailures=$failed.Count; Warnings=$warnings.Count
    }
    $sheets=@(New-LocalDiscoveryRow @{Name='Server Summary'; Rows=@($summary)})
    foreach($mapping in @(
        'System|System','Network|Network','Storage|Storage','Roles and Features|RolesAndFeatures',
        'Applications|Applications','Services|Services','Listening Ports|ListeningPorts',
        'Shares and Permissions|SharesAndPermissions','Scheduled Tasks|ScheduledTasks','Local Accounts|LocalAccounts',
        'Local Group Membership|LocalGroupMembership','Service Accounts|ServiceAccounts','Active Directory|ActiveDirectory',
        'AD Users|ADUsers','AD Group Membership|ADGroupMembership','AD Replication|ADReplication',
        'DHCP Scopes|DHCPScopes','DHCP Exclusions|DHCPExclusions','DHCP Reservations|DHCPReservations',
        'DHCP Options|DHCPOptions','DHCP Failover|DHCPFailover','DNS Zones|DNSZones','Diagnostics|Diagnostics'
    )){
        $pair=$mapping.Split('|'); $property=$Payload.PSObject.Properties[$pair[1]]; $rows=@()
        if($property){$rows=@($property.Value | Where-Object {$null -ne $_})}
        $sheets+=@(New-LocalDiscoveryRow @{Name=$pair[0]; Rows=$rows})
    }
    $parts=@{}
    $ns='http://schemas.openxmlformats.org/spreadsheetml/2006/main'
    $rels='http://schemas.openxmlformats.org/officeDocument/2006/relationships'
    $contentTypes='<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>'
    $workbook='<workbook xmlns="'+$ns+'" xmlns:r="'+$rels+'"><sheets>'
    $relationships='<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdStyles" Type="'+$rels+'/styles" Target="styles.xml"/>'
    for($index=0;$index -lt $sheets.Count;$index++){
        $sheet=$sheets[$index]; $number=$index+1; $rows=@($sheet.Rows)
        if(-not $rows.Count){$rows=@(New-LocalDiscoveryRow @{Status='No records returned. See Diagnostics for empty, skipped, or failed collectors.'})}
        if($rows.Count -gt 1048575){throw ('Sheet {0} exceeds the Excel row limit; the snapshot retains the inventory.' -f $sheet.Name)}
        $headers=New-Object 'System.Collections.Generic.List[string]'
        # Stable leading identity column, then every field seen in any row.
        if($rows[0].PSObject.Properties['ComputerName']){$headers.Add('ComputerName')}
        foreach($row in $rows){foreach($property in $row.PSObject.Properties){if(-not $headers.Contains($property.Name)){$headers.Add($property.Name)}}}
        if($headers.Count -gt 16384){throw 'Sheet exceeds the Excel column limit.'}
        $lastColumn=Get-DiscoveryColumnName $headers.Count; $lastRow=$rows.Count+1
        $xml=New-Object Text.StringBuilder
        [void]$xml.Append(('<worksheet xmlns="{0}"><dimension ref="A1:{1}{2}"/><sheetViews><sheetView workbookViewId="0" showGridLines="0"><pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen"/></sheetView></sheetViews><sheetFormatPr defaultRowHeight="30"/><cols><col min="1" max="{3}" width="26" customWidth="1"/></cols><sheetData><row r="1" ht="32" customHeight="1">' -f $ns,$lastColumn,$lastRow,$headers.Count))
        for($c=0;$c -lt $headers.Count;$c++){[void]$xml.Append((Convert-DiscoveryCellXml ((Get-DiscoveryColumnName ($c+1))+'1') $headers[$c] 1))}
        [void]$xml.Append('</row>')
        for($r=0;$r -lt $rows.Count;$r++){
            $lineCount=2
            foreach($p in $rows[$r].PSObject.Properties){
                $lines=0
                foreach($line in (([string]$p.Value) -split "`n")){$lines+=[int][math]::Max(1,[math]::Ceiling($line.Length/24.0))}
                $lineCount=[math]::Max($lineCount,$lines)
            }
            $height=[math]::Min(409,15*$lineCount)
            [void]$xml.Append(('<row r="{0}" ht="{1}" customHeight="1">' -f ($r+2),$height))
            for($c=0;$c -lt $headers.Count;$c++){
                $property=$rows[$r].PSObject.Properties[$headers[$c]]; $value=$null
                if($property){$value=$property.Value}
                $style=0
                if($headers[$c] -match 'Percent|Utilization|SizeGB|FreeGB|MemoryGB'){$style=3}
                [void]$xml.Append((Convert-DiscoveryCellXml ((Get-DiscoveryColumnName ($c+1))+($r+2)) $value $style))
            }
            [void]$xml.Append('</row>')
        }
        [void]$xml.Append(('</sheetData><autoFilter ref="A1:{0}{1}"/></worksheet>' -f $lastColumn,$lastRow))
        $parts['xl/worksheets/sheet'+$number+'.xml']=$xml.ToString()
        $workbook+='<sheet name="'+(Convert-DiscoveryXmlText $sheet.Name)+'" sheetId="'+$number+'" r:id="rId'+$number+'"/>'
        $relationships+='<Relationship Id="rId'+$number+'" Type="'+$rels+'/worksheet" Target="worksheets/sheet'+$number+'.xml"/>'
        $contentTypes+='<Override PartName="/xl/worksheets/sheet'+$number+'.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>'
    }
    $parts['xl/workbook.xml']=$workbook+'</sheets></workbook>'
    $parts['xl/_rels/workbook.xml.rels']=$relationships+'</Relationships>'
    $parts['[Content_Types].xml']=$contentTypes+'</Types>'
    $parts['_rels/.rels']='<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="'+$rels+'/officeDocument" Target="xl/workbook.xml"/></Relationships>'
    $parts['xl/styles.xml']='<styleSheet xmlns="'+$ns+'"><numFmts count="1"><numFmt numFmtId="164" formatCode="yyyy-mm-dd hh:mm:ss"/></numFmts><fonts count="2"><font><sz val="10"/><name val="Arial"/></font><font><b/><sz val="10"/><color rgb="FFFFFFFF"/><name val="Arial"/></font></fonts><fills count="3"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill><fill><patternFill patternType="solid"><fgColor rgb="FF1F4E79"/><bgColor indexed="64"/></patternFill></fill></fills><borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders><cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs><cellXfs count="4"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0" applyAlignment="1"><alignment vertical="top" wrapText="1"/></xf><xf numFmtId="0" fontId="1" fillId="2" borderId="0" xfId="0" applyFont="1" applyFill="1" applyAlignment="1"><alignment vertical="center" wrapText="1"/></xf><xf numFmtId="164" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/><xf numFmtId="2" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/></cellXfs><cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles></styleSheet>'
    [StandaloneDiscoveryZip]::Write($Path,$parts)
}
