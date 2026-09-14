function Get-LocalUserCustom {
    param (
        [string]$Username
    )
    # Get the specified local user account
    $cmdOutput = net user $Username
    if ($LASTEXITCODE -eq 0) {
        return $true  # User account exists
    } else {
        return $false  # User account does not exist
    }
}

try {
    $username = "impactlocal"
    if(!(Get-LocalUserCustom -Username $username)) {
        throw "Local user account $username not present. Remediation required."
    }
    $username = "impactadmin"
    if(!(Get-LocalUserCustom -Username $username)) {
        throw "Local user account $username not present. Remediation required."
    }
    exit 0
}
catch {
    $errMsg = $_.Exception.Message
    exit 1
}