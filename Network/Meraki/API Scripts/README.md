# Meraki Network Deployment Automation Script

## Overview

This PowerShell script automates the deployment of Meraki networks by reading configuration from a CSV file and using the Meraki Dashboard API to implement the configurations.

## Prerequisites

- PowerShell 5.1 or later
- Meraki Dashboard API key with full organization access
- CSV file containing the Meraki network deployment configuration

## Usage

```powershell
.\Deploy-MerakiNetwork.ps1 -CsvPath ".\MerakiDeployment.csv" -ApiKey "1234567890abcdef" -OutputPath ".\DeploymentResults\"
```

### Parameters

- `CsvPath`: Path to the CSV file containing the Meraki network deployment configuration.
- `ApiKey`: Meraki Dashboard API key with full organization access.
- `OutputPath`: Path to save the deployment log and results (optional, default is `.\MerakiDeploymentResults`).

## CSV Templates

The CSV file should contain sections for different configuration settings. Each section should start with a header line followed by the configuration data. Below are the templates for each section:

### Site Settings

```csv
#Site Settings
SiteName,SiteAddress,SiteCity,SiteState,SiteCountry,SitePostalCode,SiteNetworkName,SiteNetworkType,SiteNetworkTags,SiteNetworkTimeZone,SiteNetworkNotes
Site 1,123 Main St,Anytown,CA,USA,12345,Example Network 1,wireless,tag1,tag2,America/Los_Angeles,Example notes for site 1
Site 2,456 Elm St,Othertown,TX,USA,67890,Example Network 2,wireless,tag3,tag4,America/New_York,Example notes for site 2
```

### Organization Settings

```csv
#Organization Settings
OrganizationName,OrgStreetAddress,OrgCity,OrgState,OrgCountry,OrgPostalCode
Example Organization,123 Main St,Anytown,CA,USA,12345
```

### Network Settings

```csv
#Network Settings
NetworkName,NetworkType,NetworkTags,NetworkTimeZone,NetworkNotes
Example Network,wireless,tag1,tag2,America/Los_Angeles,Example notes
```

### Security Settings

```csv
#Security Settings
IPSEnabled,IPSMode,IDSRules,ContentFilteringEnabled,ContentFilteringCategories,ContentFilteringBlockedUrls,ContentFilteringAllowedUrls,AMEnabled,AMMode,AMAllowedUrls,AMBlockedUrls
true,detection,connectivity,security,example.com,example.org,true,prevention,example.com,example.org
```

### SSID Settings

```csv
#SSID Settings
SiteName,SSIDName,SSIDEnabled,SSIDAdminAccessible,SSIDAuthMode,SSIDEncryptionMode,SSIDPSK,SSIDSplashPage,SSIDWPAEncryptionMode,SSIDIPAssignmentMode,SSIDMinBitrate,SSIDVLAN,SSIDAvailableOnAllAPs,SSIDVisibility,SSIDEnterpriseAdminAccess,SSIDRadiusServers,SSIDRadiusAccountingEnabled,SSIDRadiusAttributeForGroupPolicies,SSIDRadiusOverride,SSIDRadiusCalledStationIDFormat,SSIDRadiusAuthenticationNasID,SSIDRadiusAccountingNasID,SSIDIPv6BridgeEnabled,SSIDMandatoryDHCPEnabled,SSIDRadiusProxyEnabled,SSIDWalledGardenEnabled,SSIDWalledGardenRanges
Site 1,Example SSID,true,true,psk,examplepsk,,,wpa2,,,vlan,10,12,true,true,true,192.168.1.0/24
```

### VLANs

```csv
#VLANs
SiteName,VLANId,VLANName,VLANSubnet,VLANApplianceIP,VLANDHCPEnabled,VLANDHCPLeaseTime,VLANDHCPDNSServers,VLANDHCPBootOptionsEnabled,VLANDHCPBootNextServer,VLANDHCPBootFilename,VLANReservedIPRanges,VLANDHCPOptions
Site 1,10,Example VLAN,192.168.1.0/24,192.168.1.1,true,1 day,8.8.8.8;8.8.4.4,true,192.168.1.1,bootfile,192.168.1.10:192.168.1.20:Reserved,66:ip:192.168.1.1
```

### Traffic Shaping

```csv
#Traffic Shaping
SiteName,TrafficShapingEnabled,TrafficShapingGlobalBandwidthLimits,TrafficShapingRules,TrafficShapingDSCPTaggingEnabled,TrafficShapingDSCPTags
Site 1,true,1000:1000:2000:2000,application:skype:1000:1000:high:46,true,46
```

### Firewall Rules

```csv
#Firewall Rules
SiteName,FirewallRuleComment,FirewallRulePolicy,FirewallRuleProtocol,FirewallRuleSourceCIDR,FirewallRuleSourcePort,FirewallRuleDestinationCIDR,FirewallRuleDestinationPort,FirewallRuleSyslogEnabled
Site 1,Allow all,allow,any,any,any,any,any,true
```

### Devices

```csv
#Devices
SiteName,DeviceSerial,DeviceName,DeviceNotes,DeviceTags,DeviceAddress,DeviceLatitude,DeviceLongitude,DeviceNetworkId,DeviceSwitchProfileId,DeviceFloorPlanId
Site 1,Q2XX-XXXX-XXXX,Example Device,Example notes,tag1,tag2,123 Main St,37.7749,-122.4194,N_1234,12345,67890
```

### Switch Ports

```csv
#Switch Ports
SiteName,SwitchSerial,PortNumber,PortName,PortEnabled,PortType,PortVLAN,PortVoiceVLAN,PortAllowedVLANs,PortSTpGuard,PortRSTpEnabled,PortPoEEnabled,PortIsolationEnabled,PortAccessPolicyType,PortAccessPolicyNumber,PortMacAllowList
Site 1,Q2XX-XXXX-XXXX,1,Example Port,true,access,10,20,10-20,root guard,true,true,false,MAC whitelist,1,00:11:22:33:44:55
```

### MX Settings

```csv
#MX Settings
SiteName,MXNetworkId,MXUplink1,MXUplink2,MXFailoverEnabled,MXFailoverVPNConnectivity,MXWANTrafficUplinkPreferences,MXVLANTagging,MXWarmSpareEnabled,MXWarmSpareSerial,MXCellularFailoverEnabled
Site 1,N_1234,static:192.168.1.2:192.168.1.1:255.255.255.0:8.8.8.8,8.8.4.4,true,true,application:skype:wan1,true,Q2XX-XXXX-XXXX,true
```

### VPN Settings

```csv
#VPN Settings
SiteName,VPNMode,VPNSubnets,VPNHubs,ClientVPNEnabled,ClientVPNSubnet,ClientVPNDNSServers,ClientVPNProtocol,ClientVPNSplitTunnel,ClientVPNDefaultRoute,ClientVPNPSK
Site 1,hub,192.168.1.0/24:true,12345:true,true,192.168.2.0/24,8.8.8.8,8.8.4.4,ikev2,true,true,examplepsk
```

### Alerts

```csv
#Alerts
SiteName,AlertType,AlertEnabled,AlertEmails,AlertWebhooks,AlertSNMP,AlertSeverity,AlertThreshold,AlertFrequency
Site 1,network_down,true,admin@example.com,,true,critical,80,5 minutes
```

### Group Policies

```csv
#Group Policies
SiteName,PolicyName,PolicyScheduling,PolicyBandwidthLimits,PolicyFirewallRules,PolicyVLANTagging,PolicyContentFiltering,PolicyVPNAccess
Site 1,Example Policy,monday:08:00-17:00,1000:1000,Allow all:allow:any:any:any,true
```

### Configuration Templates

```csv
#Configuration Templates
SiteName,TemplateName,TemplateNetworks,TemplateConfig
Site 1,Example Template,Office Network,config1
```

## Logging

The script generates a log file in the specified output directory. The log file contains detailed information about the deployment process, including any errors encountered.

## Author

ATimokhin

## Version

1.0

## Creation Date

March 11, 2025
