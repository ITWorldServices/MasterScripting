# Prompt for tenant interactively
$tenant = Read-Host -Prompt "Enter your Entra ID tenant domain (e.g., mytenant.onmicrosoft.com)"
$appName = "PnP.PowerShell"
$outPath = $PSScriptRoot

try {
    $result = Register-PnPEntraIDApp -ApplicationName $appName -Tenant $tenant -OutPath $outPath -DeviceLogin -ErrorAction Stop
    Write-Host "Client ID: $($result.ClientId)"
} catch {
    if ($_.Exception.Message -like "*already exists*") {
        Write-Host "App already exists. Attempting to retrieve Client ID via Microsoft Graph PowerShell..."

        # Microsoft Graph fallback. Ensure Microsoft.Graph.Applications module is installed.
        if (-not (Get-Module -ListAvailable -Name Microsoft.Graph.Applications)) {
            Write-Host "Microsoft.Graph.Applications module not found. Please install it by running 'Install-Module Microsoft.Graph.Applications'."
        } else {
            Connect-MgGraph -Scopes "Application.Read.All"
            $existing = Get-MgApplication -Filter "displayName eq '$appName'"
            if ($existing) {
                Write-Host "Client ID (existing): $($existing.AppId)"
            } else {
                Write-Host "App exists but could not retrieve Client ID automatically. Please check manually in the Entra portal."
            }
        }
    } else {
        throw
    }
}
