function Set-LocalUserPassword {
    param (
        [string]$Username,
        [string]$Password
    )
    # Set the password for the user and ensure it never expires
    $cmdOutput = net user $Username $Password /expires:never /Y
    if ($LASTEXITCODE -eq 0) {
        return $true  # Password set successfully
    } else {
        return $false  # Password set failed
    }
}
function New-LocalUserCustom {
    param (
        [string]$Username,
        [string]$Password
    )
    # Create a new local user account
    $cmdOutput = net user $Username $Password /expires:never /add /Y
    if ($LASTEXITCODE -eq 0) {
        return $true  # User account created successfully
    } else {
        return $false  # User account creation failed
    }
}
function Add-LocalGroupMemberCustom {
    param (
        [string]$Group,
        [string]$Member
    )
    # Add the specified member to the local group
    $cmdOutput = net localgroup $Group $Member /add
    if ($LASTEXITCODE -eq 0) {
        return $true  # Member added to group successfully
    } else {
        return $false  # Failed to add member to group
    }
}
function create_admin {
    param ( [string]$username, [string]$password )  # Update parameter type to securestring
    try {
        # check if the account already exists and update the password if it does
        if ((net user $Username) -match "User name") {
            if(!(Set-LocalUserPassword -Username $username -Password $password)) {
                throw "Failed to set the password"
            }
            return
        } else {
            # create the account if it doesn't exist
            if(!(New-LocalUserCustom -Username $username -Password $password)) {
                throw "Failed to create the account"
            }
            # Add the account to the administrators group
            if (!(Add-LocalGroupMemberCustom -Group "Administrators" -Member $username)) {
                throw "Failed to add the account to the administrators group"
            }
            # Set the password for the user and ensure it never expires
            if(!(Set-LocalUserPassword -Username $username -Password $password)) {
                throw "Failed to set the password"
            }
            return
        }   
    }
    catch {
        $errMsg = $_.Exception.Message
        Write-Error $errMsg
        
        exit 1
    }
}

try {
    create_admin -username "impactlocal" -password "a9MzAU-NejGnUIoxG5bpf371YJ0fYAjZNcPB3bbOpN0cje"
    create_admin -username "impactadmin" -password "Su177hkyGMiGq!"
    
    exit 0
}
catch {
    $errMsg = $_.Exception.Message
    Write-Error $errMsg
        
    exit 1
}