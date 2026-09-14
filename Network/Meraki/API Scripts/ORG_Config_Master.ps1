<#
.SYNOPSIS
    Meraki Network Deployment Automation Script
.DESCRIPTION
    This PowerShell script automates the deployment of Meraki networks by reading configuration
    from a CSV file and using the Meraki Dashboard API to implement the configurations.
.PARAMETER CsvPath
    Path to the CSV file containing the Meraki network deployment configuration.
.PARAMETER ApiKey
    Meraki Dashboard API key with full organization access.
.PARAMETER OutputPath
    Path to save the deployment log and results.
.EXAMPLE
    .\Deploy-MerakiNetwork.ps1 -CsvPath ".\MerakiDeployment.csv" -ApiKey "1234567890abcdef" -OutputPath ".\DeploymentResults\"
.NOTES
    Author: Claude
    Version: 1.0
    Creation Date: March 11, 2025
#>

param (
    [Parameter(Mandatory=$true)]
    [string]$CsvPath,
    
    [Parameter(Mandatory=$true)]
    [string]$ApiKey,
    
    [Parameter(Mandatory=$false)]
    [string]$OutputPath = ".\MerakiDeploymentResults"
)

# Create output directory if it doesn't exist
if (-not (Test-Path -Path $OutputPath)) {
    New-Item -ItemType Directory -Path $OutputPath | Out-Null
}

# Initialize logging
$logFile = Join-Path -Path $OutputPath -ChildPath "MerakiDeployment_$(Get-Date -Format 'yyyyMMdd_HHmmss').log"
function Write-Log {
    param (
        [string]$Message,
        [string]$Level = "INFO"
    )
    
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $logEntry = "[$timestamp] [$Level] $Message"
    Add-Content -Path $logFile -Value $logEntry
    
    switch ($Level) {
        "ERROR" { Write-Host $logEntry -ForegroundColor Red }
        "WARNING" { Write-Host $logEntry -ForegroundColor Yellow }
        "SUCCESS" { Write-Host $logEntry -ForegroundColor Green }
        default { Write-Host $logEntry }
    }
}

# API Helper Functions
$baseUrl = "https://api.meraki.com/api/v1"
$headers = @{
    "X-Cisco-Meraki-API-Key" = $ApiKey
    "Content-Type" = "application/json"
}

$script:apiCallCount = 0
$script:apiCallStartTime = Get-Date

function Invoke-MerakiApiRequest {
    param (
        [string]$Method,
        [string]$Endpoint,
        [object]$Body = $null,
        [int]$MaxRetries = 3,
        [int]$RetryIntervalSeconds = 1
    )
    
    # Check rate limiting
    $currentTime = Get-Date
    $timeWindow = (Get-Date).AddSeconds(-60)
    
    if ($script:apiCallStartTime -lt $timeWindow) {
        $script:apiCallCount = 0
        $script:apiCallStartTime = $currentTime
    }
    
    if ($script:apiCallCount -ge [int]$env:APIRateLimitMaxCalls) {
        $waitTime = 60 - ($currentTime - $script:apiCallStartTime).TotalSeconds
        if ($waitTime -gt 0) {
            Write-Log "Rate limit reached. Waiting $waitTime seconds..." -Level "WARNING"
            Start-Sleep -Seconds $waitTime
            $script:apiCallCount = 0
            $script:apiCallStartTime = Get-Date
        }
    }
    
    $url = "$baseUrl$Endpoint"
    $retryCount = 0
    $success = $false
    $result = $null
    
    while (-not $success -and $retryCount -lt $MaxRetries) {
        try {
            if ($Body) {
                $bodyJson = $Body | ConvertTo-Json -Depth 10
                $result = Invoke-RestMethod -Uri $url -Headers $headers -Method $Method -Body $bodyJson -ErrorAction Stop
            }
            else {
                $result = Invoke-RestMethod -Uri $url -Headers $headers -Method $Method -ErrorAction Stop
            }
            $success = $true
        }
        catch {
            $statusCode = $_.Exception.Response.StatusCode.value__
            
            # Handle rate limiting (429)
            if ($statusCode -eq 429) {
                $retryCount++
                $waitTime = [math]::Pow(2, $retryCount) * $RetryIntervalSeconds
                Write-Log "Rate limit exceeded. Waiting $waitTime seconds before retry ($retryCount/$MaxRetries)..." -Level "WARNING"
                Start-Sleep -Seconds $waitTime
            }
            # Handle server errors (5xx)
            elseif ($statusCode -ge 500 -and $statusCode -lt 600) {
                $retryCount++
                $waitTime = [math]::Pow(2, $retryCount) * $RetryIntervalSeconds
                Write-Log "Server error ($statusCode). Waiting $waitTime seconds before retry ($retryCount/$MaxRetries)..." -Level "WARNING"
                Start-Sleep -Seconds $waitTime
            }
            else {
                Write-Log "API request failed: $($_.Exception.Message)" -Level "ERROR"
                Write-Log "Request: $Method $url" -Level "ERROR"
                if ($Body) {
                    Write-Log "Request Body: $($Body | ConvertTo-Json -Depth 5)" -Level "ERROR"
                }
                throw $_
            }
        }
    }
    
    if (-not $success) {
        Write-Log "Failed after $MaxRetries retries." -Level "ERROR"
        throw "API request failed after multiple retries"
    }
    
    $script:apiCallCount++
    return $result
}

function Get-MerakiOrganizations {
    try {
        return Invoke-MerakiApiRequest -Method "GET" -Endpoint "/organizations"
    }
    catch {
        Write-Log "Failed to get organizations: $($_.Exception.Message)" -Level "ERROR"
        throw $_
    }
}

function Get-MerakiOrganization {
    param (
        [string]$OrgName
    )
    
    $orgs = Get-MerakiOrganizations
    $org = $orgs | Where-Object { $_.name -eq $OrgName }
    
    if (-not $org) {
        throw "Organization '$OrgName' not found"
    }
    
    return $org
}

function Create-MerakiOrganization {
    param (
        [string]$OrgName
    )
    
    try {
        $body = @{
            name = $OrgName
        }
        
        return Invoke-MerakiApiRequest -Method "POST" -Endpoint "/organizations" -Body $body
    }
    catch {
        Write-Log "Failed to create organization: $($_.Exception.Message)" -Level "ERROR"
        throw $_
    }
}

function Update-MerakiOrganization {
    param (
        [string]$OrgId,
        [hashtable]$OrgData
    )
    
    try {
        $endpoint = "/organizations/$OrgId"
        return Invoke-MerakiApiRequest -Method "PUT" -Endpoint $endpoint -Body $OrgData
    }
    catch {
        Write-Log "Failed to update organization: $($_.Exception.Message)" -Level "ERROR"
        throw $_
    }
}

function Get-MerakiNetworks {
    param (
        [string]$OrgId
    )
    
    try {
        $endpoint = "/organizations/$OrgId/networks"
        return Invoke-MerakiApiRequest -Method "GET" -Endpoint $endpoint
    }
    catch {
        Write-Log "Failed to get networks: $($_.Exception.Message)" -Level "ERROR"
        throw $_
    }
}

function Create-MerakiNetwork {
    param (
        [string]$OrgId,
        [string]$NetworkName,
        [string]$NetworkType,
        [string[]]$NetworkTags,
        [string]$NetworkTimeZone,
        [string]$NetworkNotes
    )
    
    try {
        $endpoint = "/organizations/$OrgId/networks"
        
        $body = @{
            name = $NetworkName
            productTypes = $NetworkType -split ','
        }
        
        if ($NetworkTags) {
            $body.tags = $NetworkTags
        }
        
        if ($NetworkTimeZone) {
            $body.timeZone = $NetworkTimeZone
        }
        
        if ($NetworkNotes) {
            $body.notes = $NetworkNotes
        }
        
        return Invoke-MerakiApiRequest -Method "POST" -Endpoint $endpoint -Body $body
    }
    catch {
        Write-Log "Failed to create network: $($_.Exception.Message)" -Level "ERROR"
        throw $_
    }
}

function Update-MerakiNetworkSettings {
    param (
        [string]$NetworkId,
        [hashtable]$NetworkSettings
    )
    
    try {
        $endpoint = "/networks/$NetworkId"
        return Invoke-MerakiApiRequest -Method "PUT" -Endpoint $endpoint -Body $NetworkSettings
    }
    catch {
        Write-Log "Failed to update network settings: $($_.Exception.Message)" -Level "ERROR"
        throw $_
    }
}

function Configure-MerakiNetworkSecurity {
    param (
        [string]$NetworkId,
        [hashtable]$SecuritySettings
    )
    
    # Configure IPS/IDS
    if ($SecuritySettings.IPSEnabled -eq "true") {
        try {
            $ipsEndpoint = "/networks/$NetworkId/security/intrusionSettings"
            $ipsBody = @{
                mode = $SecuritySettings.IPSMode
                idsRulesets = $SecuritySettings.IDSRules -split ','
            }
            
            Invoke-MerakiApiRequest -Method "PUT" -Endpoint $ipsEndpoint -Body $ipsBody
            Write-Log "Configured IPS/IDS settings for network $NetworkId" -Level "SUCCESS"
        }
        catch {
            Write-Log "Failed to configure IPS/IDS settings: $($_.Exception.Message)" -Level "ERROR"
        }
    }
    
    # Configure Content Filtering
    if ($SecuritySettings.ContentFilteringEnabled -eq "true") {
        try {
            $cfEndpoint = "/networks/$NetworkId/contentFiltering"
            $cfBody = @{}
            
            if ($SecuritySettings.ContentFilteringCategories) {
                $cfBody.blockedCategories = $SecuritySettings.ContentFilteringCategories -split ','
            }
            
            if ($SecuritySettings.ContentFilteringBlockedUrls) {
                $cfBody.blockedUrls = $SecuritySettings.ContentFilteringBlockedUrls -split ','
            }
            
            if ($SecuritySettings.ContentFilteringAllowedUrls) {
                $cfBody.allowedUrls = $SecuritySettings.ContentFilteringAllowedUrls -split ','
            }
            
            Invoke-MerakiApiRequest -Method "PUT" -Endpoint $cfEndpoint -Body $cfBody
            Write-Log "Configured Content Filtering for network $NetworkId" -Level "SUCCESS"
        }
        catch {
            Write-Log "Failed to configure Content Filtering: $($_.Exception.Message)" -Level "ERROR"
        }
    }
    
    # Configure Advanced Malware Protection (AMP)
    if ($SecuritySettings.AMEnabled -eq "true") {
        try {
            $ampEndpoint = "/networks/$NetworkId/security/malwareSettings"
            $ampBody = @{
                mode = $SecuritySettings.AMMode
            }
            
            if ($SecuritySettings.AMAllowedUrls) {
                $ampBody.allowedUrls = $SecuritySettings.AMAllowedUrls -split ','
            }
            
            if ($SecuritySettings.AMBlockedUrls) {
                $ampBody.blockedUrls = $SecuritySettings.AMBlockedUrls -split ','
            }
            
            Invoke-MerakiApiRequest -Method "PUT" -Endpoint $ampEndpoint -Body $ampBody
            Write-Log "Configured Advanced Malware Protection for network $NetworkId" -Level "SUCCESS"
        }
        catch {
            Write-Log "Failed to configure Advanced Malware Protection: $($_.Exception.Message)" -Level "ERROR"
        }
    }
}

function Create-MerakiSSID {
    param (
        [string]$NetworkId,
        [hashtable]$SSIDSettings,
        [int]$SSIDNumber
    )
    
    try {
        $endpoint = "/networks/$NetworkId/wireless/ssids/$SSIDNumber"
        
        $body = @{
            name = $SSIDSettings.SSIDName
            enabled = [System.Convert]::ToBoolean($SSIDSettings.SSIDEnabled)
            adminAccessible = [System.Convert]::ToBoolean($SSIDSettings.SSIDAdminAccessible)
            authMode = $SSIDSettings.SSIDAuthMode
            encryptionMode = $SSIDSettings.SSIDEncryptionMode
        }
        
        # Add new SSID parameters
        if ($SSIDSettings.SSIDEnterpriseAdminAccess) {
            $body.enterpriseAdminAccess = [System.Convert]::ToBoolean($SSIDSettings.SSIDEnterpriseAdminAccess)
        }
        
        if ($SSIDSettings.SSIDRadiusOverride) {
            $body.radiusOverride = [System.Convert]::ToBoolean($SSIDSettings.SSIDRadiusOverride)
        }
        
        if ($SSIDSettings.SSIDRadiusCalledStationIDFormat) {
            $body.radiusCalledStationIDFormat = $SSIDSettings.SSIDRadiusCalledStationIDFormat
        }
        
        if ($SSIDSettings.SSIDRadiusAuthenticationNasID) {
            $body.radiusAuthenticationNasID = $SSIDSettings.SSIDRadiusAuthenticationNasID
        }
        
        if ($SSIDSettings.SSIDRadiusAccountingNasID) {
            $body.radiusAccountingNasID = $SSIDSettings.SSIDRadiusAccountingNasID
        }
        
        if ($SSIDSettings.SSIDIPv6BridgeEnabled) {
            $body.ipv6BridgeEnabled = [System.Convert]::ToBoolean($SSIDSettings.SSIDIPv6BridgeEnabled)
        }
        
        if ($SSIDSettings.SSIDMandatoryDHCPEnabled) {
            $body.mandatoryDHCPEnabled = [System.Convert]::ToBoolean($SSIDSettings.SSIDMandatoryDHCPEnabled)
        }
        
        if ($SSIDSettings.SSIDRadiusProxyEnabled) {
            $body.radiusProxyEnabled = [System.Convert]::ToBoolean($SSIDSettings.SSIDRadiusProxyEnabled)
        }
        
        # Authentication settings
        if ($SSIDSettings.SSIDAuthMode) {
            $body.authMode = $SSIDSettings.SSIDAuthMode
            
            if ($SSIDSettings.SSIDAuthMode -eq "psk") {
                $body.psk = $SSIDSettings.SSIDPSK
            }
            elseif ($SSIDSettings.SSIDAuthMode -eq "open-with-splash") {
                $body.splashPage = $SSIDSettings.SSIDSplashPage
            }
            elseif ($SSIDSettings.SSIDAuthMode -eq "wpa" -or $SSIDSettings.SSIDAuthMode -eq "wpa2" -or $SSIDSettings.SSIDAuthMode -eq "wpa3") {
                if ($SSIDSettings.SSIDWPAEncryptionMode) {
                    $body.encryptionMode = $SSIDSettings.SSIDWPAEncryptionMode
                }
            }
            elseif ($SSIDSettings.SSIDAuthMode -eq "8021x-radius") {
                if ($SSIDSettings.SSIDRadiusServers) {
                    $radiusServers = @()
                    foreach ($serverConfig in ($SSIDSettings.SSIDRadiusServers -split ';')) {
                        $serverParts = $serverConfig -split ':'
                        if ($serverParts.Count -ge 3) {
                            $radiusServers += @{
                                host = $serverParts[0]
                                port = [int]$serverParts[1]
                                secret = $serverParts[2]
                            }
                        }
                    }
                    $body.radiusServers = $radiusServers
                }
                
                if ($SSIDSettings.SSIDRadiusAccountingEnabled -eq "true" -and $SSIDSettings.SSIDRadiusServers) {
                    $body.radiusAccountingEnabled = $true
                }
                
                if ($SSIDSettings.SSIDRadiusAttributeForGroupPolicies) {
                    $body.radiusAttributeForGroupPolicies = $SSIDSettings.SSIDRadiusAttributeForGroupPolicies
                }
            }
        }
        
        # IP assignment
        if ($SSIDSettings.SSIDIPAssignmentMode) {
            $body.ipAssignmentMode = $SSIDSettings.SSIDIPAssignmentMode
            
            if ($SSIDSettings.SSIDIPAssignmentMode -eq "vlan") {
                $body.defaultVlanId = [int]$SSIDSettings.SSIDVLAN
            }
        }
        
        # Advanced settings
        if ($SSIDSettings.SSIDMinBitrate) {
            $body.minBitrate = [int]$SSIDSettings.SSIDMinBitrate
        }
        
        if ($SSIDSettings.SSIDVisibility -eq "false") {
            $body.visible = $false
        }
        
        if ($SSIDSettings.SSIDAvailableOnAllAPs -eq "false") {
            $body.availableOnAllAPs = $false
        }
        
        if ($SSIDSettings.SSIDWalledGardenEnabled -eq "true" -and $SSIDSettings.SSIDWalledGardenRanges) {
            $body.walledGardenEnabled = $true
            $body.walledGardenRanges = $SSIDSettings.SSIDWalledGardenRanges -split ','
        }
        
        Invoke-MerakiApiRequest -Method "PUT" -Endpoint $endpoint -Body $body
        Write-Log "Configured SSID #$SSIDNumber '$($SSIDSettings.SSIDName)' for network $NetworkId" -Level "SUCCESS"
    }
    catch {
        Write-Log "Failed to configure SSID #$SSIDNumber: $($_.Exception.Message)" -Level "ERROR"
        throw $_
    }
}

function Create-MerakiVLAN {
    param (
        [string]$NetworkId,
        [hashtable]$VLANSettings
    )
    
    try {
        $vlanId = [int]$VLANSettings.VLANId
        $endpoint = "/networks/$NetworkId/appliance/vlans/$vlanId"
        
        $body = @{
            id = $vlanId
            name = $VLANSettings.VLANName
            subnet = $VLANSettings.VLANSubnet
            applianceIp = $VLANSettings.VLANApplianceIP
        }
        
        # DHCP settings
        if ($VLANSettings.VLANDHCPEnabled -eq "true") {
            $body.dhcpHandling = "Run a DHCP server"
            
            if ($VLANSettings.VLANDHCPLeaseTime) {
                $body.dhcpLeaseTime = $VLANSettings.VLANDHCPLeaseTime
            }
            
            if ($VLANSettings.VLANDHCPDNSServers) {
                $body.dnsNameservers = $VLANSettings.VLANDHCPDNSServers
            }
            
            if ($VLANSettings.VLANReservedIPRanges) {
                $reservedRanges = @()
                foreach ($rangeConfig in ($VLANSettings.VLANReservedIPRanges -split ';')) {
                    $rangeParts = $rangeConfig -split ':'
                    if ($rangeParts.Count -ge 3) {
                        $reservedRanges += @{
                            start = $rangeParts[0]
                            end = $rangeParts[1]
                            comment = $rangeParts[2]
                        }
                    }
                }
                $body.reservedIpRanges = $reservedRanges
            }
            
            if ($VLANSettings.VLANDHCPOptions) {
                $dhcpOptions = @()
                foreach ($optionConfig in ($VLANSettings.VLANDHCPOptions -split ';')) {
                    $optionParts = $optionConfig -split ':'
                    if ($optionParts.Count -ge 3) {
                        $dhcpOptions += @{
                            code = [int]$optionParts[0]
                            type = $optionParts[1]
                            value = $optionParts[2]
                        }
                    }
                }
                $body.dhcpOptions = $dhcpOptions
            }
            
            if ($VLANSettings.VLANDHCPBootOptionsEnabled -eq "true") {
                $body.dhcpBootOptionsEnabled = $true
                
                if ($VLANSettings.VLANDHCPBootNextServer) {
                    $body.dhcpBootNextServer = $VLANSettings.VLANDHCPBootNextServer
                }
                
                if ($VLANSettings.VLANDHCPBootFilename) {
                    $body.dhcpBootFilename = $VLANSettings.VLANDHCPBootFilename
                }
            }
        }
        elseif ($VLANSettings.VLANDHCPEnabled -eq "false") {
            $body.dhcpHandling = "Do not respond to DHCP requests"
        }
        
        Invoke-MerakiApiRequest -Method "PUT" -Endpoint $endpoint -Body $body
        Write-Log "Configured VLAN #$vlanId '$($VLANSettings.VLANName)' for network $NetworkId" -Level "SUCCESS"
    }
    catch {
        Write-Log "Failed to configure VLAN: $($_.Exception.Message)" -Level "ERROR"
        throw $_
    }
}

function Configure-MerakiFirewallRules {
    param (
        [string]$NetworkId,
        [array]$FirewallRules
    )
    
    try {
        $endpoint = "/networks/$NetworkId/appliance/firewall/l3FirewallRules"
        
        $rules = @()
        foreach ($rule in $FirewallRules) {
            $newRule = @{
                comment = $rule.FirewallRuleComment
                policy = $rule.FirewallRulePolicy
                protocol = $rule.FirewallRuleProtocol
            }
            
            if ($rule.FirewallRuleSourceCIDR) {
                $newRule.srcCidr = $rule.FirewallRuleSourceCIDR
            }
            
            if ($rule.FirewallRuleSourcePort) {
                $newRule.srcPort = $rule.FirewallRuleSourcePort
            }
            
            if ($rule.FirewallRuleDestinationCIDR) {
                $newRule.destCidr = $rule.FirewallRuleDestinationCIDR
            }
            
            if ($rule.FirewallRuleDestinationPort) {
                $newRule.destPort = $rule.FirewallRuleDestinationPort
            }
            
            if ($rule.FirewallRuleSyslogEnabled -eq "true") {
                $newRule.syslogEnabled = $true
            }
            
            $rules += $newRule
        }
        
        # Default rule (always required by Meraki)
        $rules += @{
            comment = "Default rule"
            policy = "allow"
            protocol = "any"
            srcCidr = "any"
            srcPort = "any"
            destCidr = "any"
            destPort = "any"
        }
        
        $body = @{
            rules = $rules
        }
        
        Invoke-MerakiApiRequest -Method "PUT" -Endpoint $endpoint -Body $body
        Write-Log "Configured Firewall Rules for network $NetworkId" -Level "SUCCESS"
    }
    catch {
        Write-Log "Failed to configure Firewall Rules: $($_.Exception.Message)" -Level "ERROR"
        throw $_
    }
}

function Configure-MerakiTrafficShaping {
    param (
        [string]$NetworkId,
        [hashtable]$TrafficShapingSettings
    )
    
    if ($TrafficShapingSettings.TrafficShapingEnabled -ne "true") {
        return
    }
    
    try {
        $endpoint = "/networks/$NetworkId/appliance/trafficShaping/rules"
        
        $rules = @()
        if ($TrafficShapingSettings.TrafficShapingRules) {
            foreach ($ruleConfig in ($TrafficShapingSettings.TrafficShapingRules -split ';')) {
                $ruleParts = $ruleConfig -split ':'
                if ($ruleParts.Count -ge 6) {
                    $rule = @{
                        definitions = @(
                            @{
                                type = $ruleParts[0]
                                value = $ruleParts[1]
                            }
                        )
                        perClientBandwidthLimits = @{
                            settings = "custom"
                            bandwidthLimits = @{
                                limitUp = [int]$ruleParts[2]
                                limitDown = [int]$ruleParts[3]
                            }
                        }
                        priority = $ruleParts[4]
                        dscpTagValue = [int]$ruleParts[5]
                    }
                    $rules += $rule
                }
            }
        }
        
        $body = @{
            rules = $rules
        }
        
        Invoke-MerakiApiRequest -Method "PUT" -Endpoint $endpoint -Body $body
        Write-Log "Configured Traffic Shaping Rules for network $NetworkId" -Level "SUCCESS"
        
        # Configure global bandwidth limits if specified
        if ($TrafficShapingSettings.TrafficShapingGlobalBandwidthLimits) {
            $globalLimitsEndpoint = "/networks/$NetworkId/appliance/trafficShaping/uplink/bandwidthLimits"
            $limitsParts = $TrafficShapingSettings.TrafficShapingGlobalBandwidthLimits -split ':'
            
            if ($limitsParts.Count -ge 4) {
                $limitsBody = @{
                    wan1 = @{
                        limitUp = [int]$limitsParts[0]
                        limitDown = [int]$limitsParts[1]
                    }
                    wan2 = @{
                        limitUp = [int]$limitsParts[2]
                        limitDown = [int]$limitsParts[3]
                    }
                }
                
                Invoke-MerakiApiRequest -Method "PUT" -Endpoint $globalLimitsEndpoint -Body $limitsBody
                Write-Log "Configured Global Bandwidth Limits for network $NetworkId" -Level "SUCCESS"
            }
        }
        
        # Add DSCP tagging configuration
        if ($TrafficShapingSettings.TrafficShapingDSCPTaggingEnabled -eq "true") {
            try {
                $dscpEndpoint = "/networks/$NetworkId/trafficShaping/dscpTagging/rules"
                
                $dscpRules = @()
                foreach ($tag in ($TrafficShapingSettings.TrafficShapingDSCPTags -split ',')) {
                    $dscpRules += @{
                        dscp = [int]$tag
                        description = "DSCP Tag $tag"
                    }
                }
                
                $dscpBody = @{
                    rules = $dscpRules
                }
                
                Invoke-MerakiApiRequest -Method "PUT" -Endpoint $dscpEndpoint -Body $dscpBody
                Write-Log "Configured DSCP tagging rules for network $NetworkId" -Level "SUCCESS"
            }
            catch {
                Write-Log "Failed to configure DSCP tagging: $($_.Exception.Message)" -Level "ERROR"
            }
        }
    }
    catch {
        Write-Log "Failed to configure Traffic Shaping: $($_.Exception.Message)" -Level "ERROR"
    }
}

function Configure-MerakiDevices {
    param (
        [string]$NetworkId,
        [array]$DeviceSettings
    )
    
    foreach ($device in $DeviceSettings) {
        try {
            $serial = $device.DeviceSerial
            $endpoint = "/devices/$serial"
            
            $body = @{}
            
            if ($device.DeviceName) {
                $body.name = $device.DeviceName
            }
            
            if ($device.DeviceNotes) {
                $body.notes = $device.DeviceNotes
            }
            
            if ($device.DeviceTags) {
                $body.tags = $device.DeviceTags -split ','
            }
            
            if ($device.DeviceAddress) {
                $body.address = $device.DeviceAddress
            }
            
            if ($device.DeviceLatitude -and $device.DeviceLongitude) {
                $body.lat = [double]$device.DeviceLatitude
                $body.lng = [double]$device.DeviceLongitude
            }
            
            $moveDeviceBody = @{
                serial = $serial
                networkId = $NetworkId
            }
            
            # First claim/move the device to the network if needed
            try {
                $moveEndpoint = "/networks/$NetworkId/devices/claim"
                Invoke-MerakiApiRequest -Method "POST" -Endpoint $moveEndpoint -Body $moveDeviceBody
                Write-Log "Claimed device $serial to network $NetworkId" -Level "SUCCESS"
            }
            catch {
                # Device might already be claimed, try to just update its settings
                Write-Log "Device $serial already claimed or error claiming: $($_.Exception.Message)" -Level "WARNING"
            }
            
            # Update device settings
            if ($body.Count -gt 0) {
                Invoke-MerakiApiRequest -Method "PUT" -Endpoint $endpoint -Body $body
                Write-Log "Updated device $serial settings" -Level "SUCCESS"
            }
            
            # Configure switch profiles if applicable
            if ($device.DeviceSwitchProfileId) {
                try {
                    $profileEndpoint = "/devices/$serial/switchProfileAssignment"
                    $profileBody = @{
                        profileId = $device.DeviceSwitchProfileId
                    }
                    
                    Invoke-MerakiApiRequest -Method "PUT" -Endpoint $profileEndpoint -Body $profileBody
                    Write-Log "Assigned switch profile to device $serial" -Level "SUCCESS"
                }
                catch {
                    Write-Log "Failed to assign switch profile to device $serial: $($_.Exception.Message)" -Level "ERROR"
                }
            }
            
            # Configure floor plan if applicable
            if ($device.DeviceFloorPlanId) {
                try {
                    $floorplanEndpoint = "/networks/$NetworkId/floorPlans/$($device.DeviceFloorPlanId)/devices"
                    $floorplanBody = @{
                        serial = $serial
                    }
                    
                    if ($device.DeviceLatitude -and $device.DeviceLongitude) {
                        $floorplanBody.x = [double]$device.DeviceLatitude
                        $floorplanBody.y = [double]$device.DeviceLongitude
                    }
                    
                    Invoke-MerakiApiRequest -Method "PUT" -Endpoint $floorplanEndpoint -Body $floorplanBody
                    Write-Log "Assigned device $serial to floor plan" -Level "SUCCESS"
                }
                catch {
                    Write-Log "Failed to assign device $serial to floor plan: $($_.Exception.Message)" -Level "ERROR"
                }
            }
        }
        catch {
            Write-Log "Failed to configure device $($device.DeviceSerial): $($_.Exception.Message)" -Level "ERROR"
        }
    }
}

function Configure-MerakiSwitchPorts {
    param (
        [string]$NetworkId,
        [array]$SwitchPortSettings
    )
    
    foreach ($portConfig in $SwitchPortSettings) {
        try {
            $serial = $portConfig.SwitchSerial
            $portNumber = $portConfig.PortNumber
            $endpoint = "/devices/$serial/switch/ports/$portNumber"
            
            $body = @{}
            
            if ($portConfig.PortName) {
                $body.name = $portConfig.PortName
            }
            
            if ($portConfig.PortEnabled -eq "true" -or $portConfig.PortEnabled -eq "false") {
                $body.enabled = [System.Convert]::ToBoolean($portConfig.PortEnabled)
            }
            
            if ($portConfig.PortType) {
                $body.type = $portConfig.PortType
            }
            
            if ($portConfig.PortVLAN) {
                $body.vlan = [int]$portConfig.PortVLAN
            }
            
            if ($portConfig.PortVoiceVLAN) {
                $body.voiceVlan = [int]$portConfig.PortVoiceVLAN
            }
            
            if ($portConfig.PortAllowedVLANs) {
                $body.allowedVlans = $portConfig.PortAllowedVLANs
            }
            
            if ($portConfig.PortSTpGuard) {
                $body.stpGuard = $portConfig.PortSTpGuard
            }
            
            if ($portConfig.PortRSTpEnabled -eq "true" -or $portConfig.PortRSTpEnabled -eq "false") {
                $body.rstpEnabled = [System.Convert]::ToBoolean($portConfig.PortRSTpEnabled)
            }
            
            if ($portConfig.PortPoEEnabled -eq "true" -or $portConfig.PortPoEEnabled -eq "false") {
                $body.poeEnabled = [System.Convert]::ToBoolean($portConfig.PortPoEEnabled)
            }
            
            if ($portConfig.PortIsolationEnabled -eq "true" -or $portConfig.PortIsolationEnabled -eq "false") {
                $body.isolationEnabled = [System.Convert]::ToBoolean($portConfig.PortIsolationEnabled)
            }
            
            if ($portConfig.PortAccessPolicyType) {
                $body.accessPolicyType = $portConfig.PortAccessPolicyType
                
                if ($portConfig.PortAccessPolicyNumber) {
                    $body.accessPolicyNumber = [int]$portConfig.PortAccessPolicyNumber
                }
            }
            
            if ($portConfig.PortMacAllowList) {
                $body.macAllowList = $portConfig.PortMacAllowList -split ','
            }
            
            Invoke-MerakiApiRequest -Method "PUT" -Endpoint $endpoint -Body $body
            Write-Log "Configured switch port $portNumber on device $serial" -Level "SUCCESS"
        }
        catch {
            Write-Log "Failed to configure switch port $($portConfig.PortNumber) on device $($portConfig.SwitchSerial): $($_.Exception.Message)" -Level "ERROR"
        }
    }
}

function Configure-MerakiMXSettings {
    param (
        [string]$NetworkId,
        [hashtable]$MXSettings
    )
    
    # Configure uplink settings
    if ($MXSettings.MXUplink1 -or $MXSettings.MXUplink2) {
        try {
            $uplinkEndpoint = "/networks/$NetworkId/appliance/uplinks/settings"
            
            $body = @{
                interfaces = @{
                    wan1 = @{}
                    wan2 = @{}
                }
            }
            
            if ($MXSettings.MXUplink1) {
                $wan1Settings = $MXSettings.MXUplink1 -split ':'
                if ($wan1Settings.Count -ge 3) {
                    $body.interfaces.wan1 = @{
                        enabled = $true
                        usingStaticIp = ($wan1Settings[0] -eq "static")
                    }
                    
                    if ($wan1Settings[0] -eq "static") {
                        $body.interfaces.wan1.staticIp = $wan1Settings[1]
                        $body.interfaces.wan1.staticGatewayIp = $wan1Settings[2]
                        $body.interfaces.wan1.staticSubnetMask = $wan1Settings[3]
                        
                        if ($wan1Settings.Count -ge 5) {
                            $body.interfaces.wan1.staticDns = $wan1Settings[4] -split ','
                        }
                    }
                }
            }
            
            if ($MXSettings.MXUplink2) {
                $wan2Settings = $MXSettings.MXUplink2 -split ':'
                if ($wan2Settings.Count -ge 3) {
                    $body.interfaces.wan2 = @{
                        enabled = $true
                        usingStaticIp = ($wan2Settings[0] -eq "static")
                    }
                    
                    if ($wan2Settings[0] -eq "static") {
                        $body.interfaces.wan2.staticIp = $wan2Settings[1]
                        $body.interfaces.wan2.staticGatewayIp = $wan2Settings[2]
                        $body.interfaces.wan2.staticSubnetMask = $wan2Settings[3]
                        
                        if ($wan2Settings.Count -ge 5) {
                            $body.interfaces.wan2.staticDns = $wan2Settings[4] -split ','
                        }
                    }
                }
            }
            
            Invoke-MerakiApiRequest -Method "PUT" -Endpoint $uplinkEndpoint -Body $body
            Write-Log "Configured MX uplink settings for network $NetworkId" -Level "SUCCESS"
        }
        catch {
            Write-Log "Failed to configure MX uplink settings: $($_.Exception.Message)" -Level "ERROR"
        }
    }
    
    # Configure WAN traffic preferences
    if ($MXSettings.MXWANTrafficUplinkPreferences) {
        try {
            $trafficEndpoint = "/networks/$NetworkId/appliance/traffic/uplinkPreferences"
            
            $preferences = @()
            foreach ($prefConfig in ($MXSettings.MXWANTrafficUplinkPreferences -split ';')) {
                $prefParts = $prefConfig -split ':'
                if ($prefParts.Count -ge 3) {
                    $preference = @{
                        trafficFilters = @(
                            @{
                                type = $prefParts[0]
                                value = $prefParts[1]
                            }
                        )
                        preferredUplink = $prefParts[2]
                    }
                    $preferences += $preference
                }
            }
            
            $body = @{
                preferences = $preferences
            }
            
            Invoke-MerakiApiRequest -Method "PUT" -Endpoint $trafficEndpoint -Body $body
            Write-Log "Configured MX traffic uplink preferences for network $NetworkId" -Level "SUCCESS"
        }
        catch {
            Write-Log "Failed to configure MX traffic uplink preferences: $($_.Exception.Message)" -Level "ERROR"
        }
    }
    
    # Configure warm spare
    if ($MXSettings.MXWarmSpareEnabled -eq "true" -and $MXSettings.MXWarmSpareSerial) {
        try {
            $spareEndpoint = "/networks/$NetworkId/appliance/warmSpare"
            
            $body = @{
                enabled = $true
                spareSerial = $MXSettings.MXWarmSpareSerial
            }
            
            Invoke-MerakiApiRequest -Method "PUT" -Endpoint $spareEndpoint -Body $body
            Write-Log "Configured MX warm spare for network $NetworkId" -Level "SUCCESS"
        }
        catch {
            Write-Log "Failed to configure MX warm spare: $($_.Exception.Message)" -Level "ERROR"
        }
    }
    
    # Configure cellular failover
    if ($MXSettings.MXCellularFailoverEnabled -eq "true") {
        try {
            $cellularEndpoint = "/networks/$NetworkId/appliance/cellular/failover"
            
            $body = @{
                enabled = $true
            }
            
            Invoke-MerakiApiRequest -Method "PUT" -Endpoint $cellularEndpoint -Body $body
            Write-Log "Configured MX cellular failover for network $NetworkId" -Level "SUCCESS"
        }
        catch {
            Write-Log "Failed to configure MX cellular failover: $($_.Exception.Message)" -Level "ERROR"
        }
    }
    
    # Configure SD-WAN settings
    if ($MXSettings.SDWANEnabled -eq "true") {
        try {
            $sdwanEndpoint = "/networks/$NetworkId/appliance/sdwan"
            
            $body = @{
                enabled = $true
                uplinks = @()
            }
            
            if ($MXSettings.SDWANUplink1) {
                $uplink1 = Parse-UplinkSettings -UplinkString $MXSettings.SDWANUplink1
                $body.uplinks += $uplink1
            }
            
            if ($MXSettings.SDWANUplink2) {
                $uplink2 = Parse-UplinkSettings -UplinkString $MXSettings.SDWANUplink2
                $body.uplinks += $uplink2
            }
            
            if ($MXSettings.SDWANPathPreferences) {
                $body.pathPreferences = Parse-PathPreferences -PreferencesString $MXSettings.SDWANPathPreferences
            }
            
            if ($MXSettings.SDWANVPNExclusions) {
                $body.vpnExclusions = $MXSettings.SDWANVPNExclusions -split ','
            }
            
            Invoke-MerakiApiRequest -Method "PUT" -Endpoint $sdwanEndpoint -Body $body
            Write-Log "Configured SD-WAN settings for network $NetworkId" -Level "SUCCESS"
        }
        catch {
            Write-Log "Failed to configure SD-WAN settings: $($_.Exception.Message)" -Level "ERROR"
        }
    }
}

function Parse-UplinkSettings {
    param (
        [string]$UplinkString
    )
    
    $settings = $UplinkString -split ':'
    $uplink = @{
        interface = $settings[0]
    }
    
    if ($settings[0] -eq "static") {
        $uplink.staticIp = $settings[1]
        $uplink.staticGateway = $settings[2]
        $uplink.staticSubnetMask = $settings[3]
    }
    
    return $uplink
}

function Parse-PathPreferences {
    param (
        [string]$PreferencesString
    )
    
    $preferences = @()
    foreach ($pref in ($PreferencesString -split ';')) {
        $prefParts = $pref -split ':'
        $preferences += @{
            type = $prefParts[0]
            value = $prefParts[1]
            preferredUplink = $prefParts[2]
        }
    }
    
    return $preferences
}

function Configure-MerakiVPN {
    param (
        [string]$NetworkId,
        [hashtable]$VPNSettings
    )
    
    # Configure site-to-site VPN
    if ($VPNSettings.VPNMode) {
        try {
            $vpnEndpoint = "/networks/$NetworkId/appliance/vpn/siteToSiteVpn"
            
            $body = @{
                mode = $VPNSettings.VPNMode
            }
            
            if ($VPNSettings.VPNMode -eq "hub" -or $VPNSettings.VPNMode -eq "spoke") {
                $subnets = @()
                
                foreach ($subnetConfig in ($VPNSettings.VPNSubnets -split ';')) {
                    $subnetParts = $subnetConfig -split ':'
                    if ($subnetParts.Count -ge 2) {
                        $subnet = @{
                            localSubnet = $subnetParts[0]
                            useVpn = [System.Convert]::ToBoolean($subnetParts[1])
                        }
                        $subnets += $subnet
                    }
                }
                
                $body.subnets = $subnets
                
                if ($VPNSettings.VPNMode -eq "hub" -and $VPNSettings.VPNHubs) {
                    $body.hubs = @()
                    foreach ($hubConfig in ($VPNSettings.VPNHubs -split ';')) {
                        $hubParts = $hubConfig -split ':'
                        if ($hubParts.Count -ge 2) {
                            $hub = @{
                                hubId = $hubParts[0]
                                useDefaultRoute = [System.Convert]::ToBoolean($hubParts[1])
                            }
                            $body.hubs += $hub
                        }
                    }
                }
            }
            
            Invoke-MerakiApiRequest -Method "PUT" -Endpoint $vpnEndpoint -Body $body
            Write-Log "Configured site-to-site VPN for network $NetworkId" -Level "SUCCESS"
        }
        catch {
            Write-Log "Failed to configure site-to-site VPN: $($_.Exception.Message)" -Level "ERROR"
        }
    }
    
    # Configure client VPN
    if ($VPNSettings.ClientVPNEnabled -eq "true") {
        try {
            $clientVpnEndpoint = "/networks/$NetworkId/appliance/vpn/clientVpn"
            
            $body = @{
                enabled = $true
            }
            
            if ($VPNSettings.ClientVPNSubnet) {
                $body.subnet = $VPNSettings.ClientVPNSubnet
            }
            
            if ($VPNSettings.ClientVPNDNSServers) {
                $body.dnsServers = $VPNSettings.ClientVPNDNSServers -split ','
            }
            
            if ($VPNSettings.ClientVPNProtocol) {
                $body.vpnMode = $VPNSettings.ClientVPNProtocol
            }
            
            if ($VPNSettings.ClientVPNSplitTunnel -eq "true") {
                $body.splitTunnel = $true
            }
            
            if ($VPNSettings.ClientVPNDefaultRoute -eq "true") {
                $body.defaultRoute = $true
            }
            
            if ($VPNSettings.ClientVPNPSK) {
                $body.psk = $VPNSettings.ClientVPNPSK
            }
            
            Invoke-MerakiApiRequest -Method "PUT" -Endpoint $clientVpnEndpoint -Body $body
            Write-Log "Configured client VPN for network $NetworkId" -Level "SUCCESS"
        }
        catch {
            Write-Log "Failed to configure client VPN: $($_.Exception.Message)" -Level "ERROR"
        }
    }
}

function Configure-MerakiAlerts {
    param (
        [string]$OrgId,
        [array]$AlertSettings
    )
    
    try {
        $endpoint = "/organizations/$OrgId/alerts/profiles"
        
        $existingProfiles = Invoke-MerakiApiRequest -Method "GET" -Endpoint $endpoint
        
        # Group alerts by type to create profiles
        $alertsByType = $AlertSettings | Group-Object -Property AlertType
        
        foreach ($alertGroup in $alertsByType) {
            $alertType = $alertGroup.Name
            $alertsOfType = $alertGroup.Group
            
            # Check if profile exists
            $existingProfile = $existingProfiles | Where-Object { $_.type -eq $alertType }
            
            # Create alert configuration
            $alertObjects = @()
            foreach ($alert in $alertsOfType) {
                if ($alert.AlertEnabled -ne "true") {
                    continue
                }
                
                $alertObj = @{
                    type = $alert.AlertType
                    enabled = $true
                    alertDestinations = @{
                        emails = @()
                        allAdmins = $false
                        snmp = $false
                        webhooks = @()
                    }
                }
                
                if ($alert.AlertEmails) {
                    $alertObj.alertDestinations.emails = $alert.AlertEmails -split ','
                }
                
                if ($alert.AlertWebhooks) {
                    $alertObj.alertDestinations.webhooks = $alert.AlertWebhooks -split ','
                }
                
                if ($alert.AlertSNMP -eq "true") {
                    $alertObj.alertDestinations.snmp = $true
                }
                
                if ($alert.AlertSeverity) {
                    $alertObj.alertSeverity = $alert.AlertSeverity
                }
                
                if ($alert.AlertThreshold) {
                    $alertObj.threshold = @{
                        value = [int]$alert.AlertThreshold
                        unit = "percentage"
                    }
                }
                
                if ($alert.AlertFrequency) {
                    $alertObj.frequency = $alert.AlertFrequency
                }
                
                $alertObjects += $alertObj
            }
            
            if ($alertObjects.Count -eq 0) {
                continue
            }
            
            # Create or update profile
            if ($existingProfile) {
                $updateEndpoint = "$endpoint/$($existingProfile.id)"
                $body = @{
                    enabled = $true
                    type = $alertType
                    alerts = $alertObjects
                }
                
                Invoke-MerakiApiRequest -Method "PUT" -Endpoint $updateEndpoint -Body $body
                Write-Log "Updated alert profile for type $alertType" -Level "SUCCESS"
            }
            else {
                $body = @{
                    name = "Generated profile for $alertType"
                    enabled = $true
                    type = $alertType
                    alerts = $alertObjects
                }
                
                Invoke-MerakiApiRequest -Method "POST" -Endpoint $endpoint -Body $body
                Write-Log "Created new alert profile for type $alertType" -Level "SUCCESS"
            }
        }
    }
    catch {
        Write-Log "Failed to configure alerts: $($_.Exception.Message)" -Level "ERROR"
    }
}

function Configure-MerakiGroupPolicies {
    param (
        [string]$NetworkId,
        [array]$PolicySettings
    )
    
    foreach ($policy in $PolicySettings) {
        try {
            $policyName = $policy.PolicyName
            
            # Check if policy exists
            $policiesEndpoint = "/networks/$NetworkId/groupPolicies"
            $existingPolicies = Invoke-MerakiApiRequest -Method "GET" -Endpoint $policiesEndpoint
            $existingPolicy = $existingPolicies | Where-Object { $_.name -eq $policyName }
            
            $body = @{
                name = $policyName
            }
            
            # Scheduling
            if ($policy.PolicyScheduling) {
                $scheduleParts = $policy.PolicyScheduling -split ':'
                if ($scheduleParts.Count -ge 2) {
                    $body.scheduling = @{
                        enabled = $true
                        monday = @{
                            active = ($scheduleParts[0] -like "*monday*")
                            from = $scheduleParts[1].Split('-')[0]
                            to = $scheduleParts[1].Split('-')[1]
                        }
                        tuesday = @{
                            active = ($scheduleParts[0] -like "*tuesday*")
                            from = $scheduleParts[1].Split('-')[0]
                            to = $scheduleParts[1].Split('-')[1]
                        }
                        wednesday = @{
                            active = ($scheduleParts[0] -like "*wednesday*")
                            from = $scheduleParts[1].Split('-')[0]
                            to = $scheduleParts[1].Split('-')[1]
                        }
                        thursday = @{
                            active = ($scheduleParts[0] -like "*thursday*")
                            from = $scheduleParts[1].Split('-')[0]
                            to = $scheduleParts[1].Split('-')[1]
                        }
                        friday = @{
                            active = ($scheduleParts[0] -like "*friday*")
                            from = $scheduleParts[1].Split('-')[0]
                            to = $scheduleParts[1].Split('-')[1]
                        }
                        saturday = @{
                            active = ($scheduleParts[0] -like "*saturday*")
                            from = $scheduleParts[1].Split('-')[0]
                            to = $scheduleParts[1].Split('-')[1]
                        }
                        sunday = @{
                            active = ($scheduleParts[0] -like "*sunday*")
                            from = $scheduleParts[1].Split('-')[0]
                            to = $scheduleParts[1].Split('-')[1]
                        }
                    }
                }
            }
            
            # Bandwidth limits
            if ($policy.PolicyBandwidthLimits) {
                $limitsParts = $policy.PolicyBandwidthLimits -split ':'
                if ($limitsParts.Count -ge 2) {
                    $body.bandwidth = @{
                        settings = "custom"
                        limits = @{
                            limitUp = [int]$limitsParts[0]
                            limitDown = [int]$limitsParts[1]
                        }
                    }
                }
            }
            
            # Firewall rules
            if ($policy.PolicyFirewallRules) {
                $firewallRules = @()
                foreach ($ruleConfig in ($policy.PolicyFirewallRules -split ';')) {
                    $ruleParts = $ruleConfig -split ':'
                    if ($ruleParts.Count -ge 5) {
                        $rule = @{
                            comment = $ruleParts[0]
                            policy = $ruleParts[1]
                            protocol = $ruleParts[2]
                            destPort = $ruleParts[3]
                            destCidr = $ruleParts[4]
                        }
                        $firewallRules += $rule
                    }
                }
                
                $body.firewall = @{
                    settings = "custom"
                    rules = $firewallRules
                }
            }
            
            # VLAN tagging
            if ($policy.PolicyVLANTagging) {
                $vlanParts = $policy.PolicyVLANTagging -split ':'
                if ($vlanParts.Count -ge 1) {
                    $body.vlanTagging = @{
                        settings = "custom"
                        vlanId = [int]$vlanParts[0]
                    }
                }
            }
            
            # Content filtering
            if ($policy.PolicyContentFiltering) {
                $cfParts = $policy.PolicyContentFiltering -split ':'
                if ($cfParts.Count -ge 1) {
                    $body.contentFiltering = @{
                        settings = "custom"
                    }
                    
                    if ($cfParts[0] -ne "none") {
                        $body.contentFiltering.allowedUrlPatterns = @{
                            patterns = $cfParts[0] -split ','
                        }
                    }
                    
                    if ($cfParts.Count -ge 2 -and $cfParts[1] -ne "none") {
                        $body.contentFiltering.blockedUrlPatterns = @{
                            patterns = $cfParts[1] -split ','
                        }
                    }
                    
                    if ($cfParts.Count -ge 3 -and $cfParts[2] -ne "none") {
                        $body.contentFiltering.blockedUrlCategories = @{
                            categories = $cfParts[2] -split ','
                        }
                    }
                }
            }
            
            # VPN access
            if ($policy.PolicyVPNAccess -eq "true" -or $policy.PolicyVPNAccess -eq "false") {
                $body.vpn = @{
                    settings = "custom"
                    client = @{
                        accessible = [System.Convert]::ToBoolean($policy.PolicyVPNAccess)
                    }
                }
            }
            
            # Create or update policy
            if ($existingPolicy) {
                $updateEndpoint = "$policiesEndpoint/$($existingPolicy.id)"
                Invoke-MerakiApiRequest -Method "PUT" -Endpoint $updateEndpoint -Body $body
                Write-Log "Updated group policy: $policyName" -Level "SUCCESS"
            }
            else {
                Invoke-MerakiApiRequest -Method "POST" -Endpoint $policiesEndpoint -Body $body
                Write-Log "Created new group policy: $policyName" -Level "SUCCESS"
            }
        }
        catch {
            Write-Log "Failed to configure group policy $($policy.PolicyName): $($_.Exception.Message)" -Level "ERROR"
        }
    }
}

function Process-MerakiDeployment {
    param (
        [string]$CsvPath
    )
    
    Write-Log "Starting Meraki deployment from CSV: $CsvPath" -Level "INFO"
    
    # Read and parse CSV
    $rawCsvContent = Get-Content -Path $CsvPath -Raw
    $sections = $rawCsvContent -split '\r?\n#'
    
    # Parse each section
    $orgSettings = @{}
    $networkSettings = @{}
    $securitySettings = @{}
    $ssidSettings = @()
    $vlanSettings = @()
    $trafficShapingSettings = @{}
    $firewallRules = @()
    $deviceSettings = @()
    $switchPortSettings = @()
    $mxSettings = @{}
    $vpnSettings = @{}
    $alertSettings = @()
    $policySettings = @()
    $templateSettings = @()
    $siteSettings = @()
    
    foreach ($section in $sections) {
        $sectionLines = $section -split '\r?\n'
        $sectionName = $sectionLines[0].Trim()
        $sectionData = $sectionLines | Where-Object { -not [string]::IsNullOrWhiteSpace($_) -and -not $_.StartsWith('#') }
        
        if ($sectionData.Count -lt 2) {
            continue
        }
        
        $headers = $sectionData[0].Split(',')
        $values = $sectionData[1].Split(',')
        
        $sectionDataObj = @{}
        for ($i = 0; $i -lt [Math]::Min($headers.Count, $values.Count); $i++) {
            $sectionDataObj[$headers[$i]] = $values[$i]
        }
        
        # Special case for multi-row sections
        if ($sectionData.Count -gt 2) {
            $multiRowData = @()
            for ($rowIndex = 1; $rowIndex -lt $sectionData.Count; $rowIndex++) {
                $rowValues = $sectionData[$rowIndex].Split(',')
                $rowObj = @{}
                for ($i = 0; $i -lt [Math]::Min($headers.Count, $rowValues.Count); $i++) {
                    $rowObj[$headers[$i]] = $rowValues[$i]
                }
                $multiRowData += $rowObj
            }
        }
        
        switch -Wildcard ($sectionName) {
            "*Organization Settings*" {
                $orgSettings = $sectionDataObj
            }
            "*Network Settings*" {
                $networkSettings = $sectionDataObj
            }
            "*Security Settings*" {
                $securitySettings = $sectionDataObj
            }
            "*SSID Settings*" {
                $ssidSettings = $multiRowData
            }
            "*VLANs*" {
                $vlanSettings = $multiRowData
            }
            "*Traffic Shaping*" {
                $trafficShapingSettings = $sectionDataObj
            }
            "*Firewall Rules*" {
                $firewallRules = $multiRowData
            }
            "*Devices*" {
                $deviceSettings = $multiRowData
            }
            "*Switch Ports*" {
                $switchPortSettings = $multiRowData
            }
            "*MX Settings*" {
                $mxSettings = $sectionDataObj
            }
            "*SD-WAN Settings*" {
                $sdwanSettings = $sectionDataObj
                # Combine SD-WAN settings with MX settings
                foreach ($key in $sdwanSettings.Keys) {
                    $mxSettings[$key] = $sdwanSettings[$key]
                }
            }
            "*Site-to-Site VPN*" -or "*AutoVPN*" -or "*Client VPN*" {
                foreach ($key in $sectionDataObj.Keys) {
                    $vpnSettings[$key] = $sectionDataObj[$key]
                }
            }
            "*Alerts*" {
                $alertSettings = $multiRowData
            }
            "*Group Policies*" {
                $policySettings = $multiRowData
            }
            "*Configuration Templates*" {
                $templateSettings = $multiRowData
            }
            "*Site Settings*" {
                $siteSettings += $sectionDataObj
            }
        }
    }
    
    # Begin deployment
    try {
        # Step 1: Get or create organization
        $orgName = $orgSettings.OrganizationName
        Write-Log "Processing organization: $orgName" -Level "INFO"
        
        try {
            $org = Get-MerakiOrganization -OrgName $orgName
            Write-Log "Found existing organization: $orgName (ID: $($org.id))" -Level "INFO"
        }
        catch {
            Write-Log "Organization not found, attempting to create: $orgName" -Level "INFO"
            $org = Create-MerakiOrganization -OrgName $orgName
            Write-Log "Created new organization: $orgName (ID: $($org.id))" -Level "SUCCESS"
        }
        
        $orgId = $org.id
        
        # Step 2: Update organization settings
        $orgUpdateData = @{}
        
        if ($orgSettings.OrgStreetAddress -or $orgSettings.OrgCity -or $orgSettings.OrgState -or $orgSettings.OrgCountry -or $orgSettings.OrgPostalCode) {
            $orgUpdateData.address = @{}
            
            if ($orgSettings.OrgStreetAddress) {
                $orgUpdateData.address.street = $orgSettings.OrgStreetAddress
            }
            
            if ($orgSettings.OrgCity) {
                $orgUpdateData.address.city = $orgSettings.OrgCity
            }
            
            if ($orgSettings.OrgState) {
                $orgUpdateData.address.state = $orgSettings.OrgState
            }
            
            if ($orgSettings.OrgCountry) {
                $orgUpdateData.address.country = $orgSettings.OrgCountry
            }
            
            if ($orgSettings.OrgPostalCode) {
                $orgUpdateData.address.postalCode = $orgSettings.OrgPostalCode
            }
        }
        
        if ($orgUpdateData.Count -gt 0) {
            Update-MerakiOrganization -OrgId $orgId -OrgData $orgUpdateData
            Write-Log "Updated organization settings for $orgName" -Level "SUCCESS"
        }
        
        # Step 3: Process each site
        foreach ($site in $siteSettings) {
            $siteName = $site.SiteName
            Write-Log "Processing site: $siteName" -Level "INFO"
            
            # Step 4: Get or create network for the site
            $networkName = $site.SiteNetworkName
            Write-Log "Processing network: $networkName" -Level "INFO"
            
            $networks = Get-MerakiNetworks -OrgId $orgId
            $network = $networks | Where-Object { $_.name -eq $networkName }
            
            if (-not $network) {
                Write-Log "Network not found, attempting to create: $networkName" -Level "INFO"
                $network = Create-MerakiNetwork -OrgId $orgId -NetworkName $networkName -NetworkType $site.SiteNetworkType -NetworkTags $site.SiteNetworkTags -NetworkTimeZone $site.SiteNetworkTimeZone -NetworkNotes $site.SiteNetworkNotes
                Write-Log "Created new network: $networkName (ID: $($network.id))" -Level "SUCCESS"
            }
            else {
                Write-Log "Found existing network: $networkName (ID: $($network.id))" -Level "INFO"
            }
            
            $networkId = $network.id
            
            # Step 5: Update network settings
            $networkUpdateData = @{}
            
            if ($site.SiteNetworkNotes) {
                $networkUpdateData.notes = $site.SiteNetworkNotes
            }
            
            if ($networkUpdateData.Count -gt 0) {
                Update-MerakiNetworkSettings -NetworkId $networkId -NetworkSettings $networkUpdateData
                Write-Log "Updated network settings for $networkName" -Level "SUCCESS"
            }
            
            # Step 6: Configure network security settings
            Configure-MerakiNetworkSecurity -NetworkId $networkId -SecuritySettings $securitySettings
            
            # Step 7: Configure SSIDs
            $siteSSIDSettings = $ssidSettings | Where-Object { $_.SiteName -eq $siteName }
            $ssidNumber = 0
            foreach ($ssid in $siteSSIDSettings) {
                Create-MerakiSSID -NetworkId $networkId -SSIDSettings $ssid -SSIDNumber $ssidNumber
                $ssidNumber++
            }
            
            # Step 8: Configure VLANs
            $siteVLANSettings = $vlanSettings | Where-Object { $_.SiteName -eq $siteName }
            foreach ($vlan in $siteVLANSettings) {
                Create-MerakiVLAN -NetworkId $networkId -VLANSettings $vlan
            }
            
            # Step 9: Configure traffic shaping
            $siteTrafficShapingSettings = $trafficShapingSettings | Where-Object { $_.SiteName -eq $siteName }
            Configure-MerakiTrafficShaping -NetworkId $networkId -TrafficShapingSettings $siteTrafficShapingSettings
            
            # Step 10: Configure firewall rules
            $siteFirewallRules = $firewallRules | Where-Object { $_.SiteName -eq $siteName }
            Configure-MerakiFirewallRules -NetworkId $networkId -FirewallRules $siteFirewallRules
            
            # Step 11: Configure devices
            $siteDeviceSettings = $deviceSettings | Where-Object { $_.SiteName -eq $siteName }
            Configure-MerakiDevices -NetworkId $networkId -DeviceSettings $siteDeviceSettings
            
            # Step 12: Configure switch ports
            $siteSwitchPortSettings = $switchPortSettings | Where-Object { $_.SiteName -eq $siteName }
            Configure-MerakiSwitchPorts -NetworkId $networkId -SwitchPortSettings $siteSwitchPortSettings
            
            # Step 13: Configure MX settings
            $siteMXSettings = $mxSettings | Where-Object { $_.SiteName -eq $siteName }
            Configure-MerakiMXSettings -NetworkId $networkId -MXSettings $siteMXSettings
            
            # Step 14: Configure VPN settings
            $siteVPNSettings = $vpnSettings | Where-Object { $_.SiteName -eq $siteName }
            Configure-MerakiVPN -NetworkId $networkId -VPNSettings $siteVPNSettings
            
            # Step 15: Configure alerts
            $siteAlertSettings = $alertSettings | Where-Object { $_.SiteName -eq $siteName }
            Configure-MerakiAlerts -OrgId $orgId -AlertSettings $siteAlertSettings
            
            # Step 16: Configure group policies
            $sitePolicySettings = $policySettings | Where-Object { $_.SiteName -eq $siteName }
            Configure-MerakiGroupPolicies -NetworkId $networkId -PolicySettings $sitePolicySettings
        }
        
        Write-Log "Meraki deployment completed successfully" -Level "SUCCESS"
    }
    catch {
        Write-Log "Meraki deployment failed: $($_.Exception.Message)" -Level "ERROR"
    }
}

# Start the deployment process
Process-MerakiDeployment -CsvPath $CsvPath