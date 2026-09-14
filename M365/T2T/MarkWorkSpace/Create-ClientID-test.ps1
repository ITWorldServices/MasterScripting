try {
    $appRegistration = Register-PnPAzureADApp -ApplicationName "PnP.PowerShell" -Tenant "506brewer.onmicrosoft.com" -Store CurrentUser
    $clientId = $appRegistration."AzureAppId/ClientId"
    Write-Host "Your Client ID is: $clientId"
}
catch {
    Write-Host "App is already registered. Please navigate to https://portal.azure.com,"
    Write-Host "Go to Azure Active Directory > App registrations, locate the 'PnP.PowerShell' app by name,"
    Write-Host "and copy the Client ID (Application ID) from the Overview page."
}
