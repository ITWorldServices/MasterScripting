<#
.SYNOPSIS
Get-WIFIPassword - Retrieve the passwords of stored Wi-Fi networks using netsh.

.DESCRIPTION 
This PowerShell function retrieves the passwords of stored Wi-Fi networks by utilizing the netsh command-line tool. It helps users access the Wi-Fi passwords for networks they have connected to previously.

.EXAMPLE
Get-WIFIPassword
Retrieves and displays the Wi-Fi passwords of all stored networks.

.INPUTS
None. You cannot pipe input to this function.

.OUTPUTS
The function returns a collection of custom objects containing the ProfileName (Wi-Fi network name) and Password (Wi-Fi password) properties.

.NOTES
- This function requires administrative privileges as it uses the netsh command-line tool.
- The Wi-Fi passwords retrieved are only applicable to the currently logged-in user's profile.

.COMPONENT
This function is part of a PowerShell module or script that can be used to manage Wi-Fi networks.

.LINK
Script Repository: <Link to the repository where the script/module is hosted>

#>

[CmdletBinding()]
param (
    # No parameters defined
)

# Get the list of stored Wi-Fi profiles using netsh, extract profile names, and retrieve their passwords
netsh wlan show profile |
    where {$_ -match '\:\s(.+)$'} |
        foreach {
            $name = $Matches[1]
            netsh wlan show profile name="$name" key=clear |
                where {$_ -match 'Key Content\W+\:\s(.+)$'} |
                    foreach {
                        [PSCustomObject]@{
                            ProfileName = $name
                            Password = $Matches[1]
                        }
                    }
        }