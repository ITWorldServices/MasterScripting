#Connect-ExchangeOnline -ExchangeEnvironmentName O365USGovGCCHigh
# OR
#Connect-ExchangeOnline 

Start-Transcript GetAcceptedDomain.txt -NoClobber
$FormatEnumerationLimit=-1
Get-AcceptedDomain | FL
Stop-Transcript

Disconnect-ExchangeOnline