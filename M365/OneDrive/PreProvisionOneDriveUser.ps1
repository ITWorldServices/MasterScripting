# Connect to M365 Sharepoint with the proper credentials. You will need to modify the line below.
Connect-SPOService -Url https://icsholdingllc-admin.sharepoint.com -Credential impactadmin@brifutelectric.com
# This is a one off script but can be looped.
Request-SPOPersonalSite -UserEmails accubid@icsholdingllc.onmicrosoft.com
Disconnect-SPOService