########################################################################
#
# The script will run all of the commands found in the Exchange on-Prem
# to O365 discovery spreadsheet
#
########################################################################

#Download the Exchange Online module if you have not done so before
#Copy the line Below to install the module. The line is currently commented out

# Install-Module -Name ExchangeOnlineManagement -AllowClobber -Scope CurrentUser
# Install-Module -Name ExchangePowerShell

#Connect to ExchangeOnlice
#Import-Module ExchangeOnlineManagement
#Connect-ExchangeOnline


# Check if Active Directory module is installed
$adModuleInstalled = Get-Module -Name ActiveDirectory -ListAvailable

if (!$adModuleInstalled) {
    Write-Host "Active Directory module is not installed. Installing..."
    Install-WindowsFeature RSAT-AD-PowerShell
} else {
    Write-Host "Active Directory module is already installed."
}


#Install the Excel Plug-in. We want to be able to export to Excel.
# Check if ImportExcel module is installed
$importExcelModuleInstalled = Get-Module -Name ImportExcel -ListAvailable

if (!$importExcelModuleInstalled) {
    Write-Host "ImportExcel module is not installed. Installing..."
    Install-Module -Name ImportExcel -Scope CurrentUser -Force
} else {
    Write-Host "ImportExcel module is already installed."
}
Install-Module -Name ImportExcel -Scope CurrentUser -Force

# Import the required modules
Import-Module ActiveDirectory
Import-Module ImportExcel


#Now get connected to Exchange on Prem to get data
$UserCredential = Get-Credential
$Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri http://exch-2010.idealease.inc/PowerShell/ -Authentication Kerberos -Credential $UserCredential
Import-PSSession $Session -DisableNameChecking

###############################################################################################################################################
###############################################################################################################################################
#Create a storage folder for the Excel document
New-Item "C:\Impact" -itemtype Directory

###############################################################################################################################################
###############################################################################################################################################
#First get a list of AD Users for lookup purposes. This is the ADUsers tab
#Get-ADUser -Properties "*" -Filter "*" | Export-CSV -Path "c:\Impact\ADUsers.csv"
Get-ADUser -Properties "*" -Filter "*" | Export-Excel C:\Impact\Exchange_O365.xlsx -Append -WorksheetName "ADUsers"

###############################################################################################################################################
###############################################################################################################################################
#Get a list of Contacts
#Get-ADObject -Filter {ObjectClass -eq 'contact'} -Properties * | Export-CSV -Path "C:\Impact\contacts.csv" -NoTypeInformation
Get-ADObject -Filter {ObjectClass -eq 'contact'} -Properties * | Export-Excel C:\Impact\Exchange_O365.xlsx -Append -WorksheetName "Contacts" 

###############################################################################################################################################
###############################################################################################################################################
#Get Resources and Rooms
#Get-Mailbox -RecipientTypeDetails RoomMailbox,EquipmentMailbox | Select-Object DisplayName,PrimarySmtpAddress | Export-Csv -Path "C:\Impact\resource_mailboxes.csv" -NoTypeInformation
Get-Mailbox -RecipientTypeDetails RoomMailbox,EquipmentMailbox | Select-Object DisplayName,PrimarySmtpAddress| Export-Excel C:\Impact\Exchange_O365.xlsx -Append -WorksheetName "Resources" 

###############################################################################################################################################
###############################################################################################################################################
#Get a list of all mailboxes
#Get-Mailbox -Filter "*" | Export-Csv -Path "c:\impact\allmailboxes.csv"
Get-Mailbox | Export-Excel C:\Impact\Exchange_O365.xlsx -Append -WorksheetName "AllMailboxes" 

###############################################################################################################################################
###############################################################################################################################################
#Get the mailbox statistics
#Get-Mailbox -Filter "*" | Get-MailboxStatistics -ErrorAction Continue | Export-Csv c:\impact\mailbox_statistics.csv
Get-Mailbox | Get-MailboxStatistics -ErrorAction Continue | Export-Excel C:\Impact\Exchange_O365.xlsx -Append -WorksheetName "Mailbox_Statistics" 

###############################################################################################################################################
###############################################################################################################################################
#Get Public Folders
#Get-PublicFolder -Recurse –ResultSize Unlimited | where{$_.MailEnabled -eq $true}  | Export-CSV c:\impact\publicfolders.csv
Get-PublicFolder -Recurse –ResultSize Unlimited | where{$_.MailEnabled -eq $true}  | Export-Excel C:\Impact\Exchange_O365.xlsx -Append -WorksheetName "Public_Folders"

###############################################################################################################################################
###############################################################################################################################################
#Get the Public folder Premissions
#Get-PublicFolder -Recurse -ResultSize Unlimited | Get-PublicFolderClientPermission | Export-Csv c:\impact\pubfoldperm.csv
Get-PublicFolder -Recurse -ResultSize Unlimited | Get-PublicFolderClientPermission | Export-Excel C:\Impact\Exchange_O365.xlsx -Append -WorksheetName "Public_Folder_Perms"

###############################################################################################################################################
###############################################################################################################################################
#Get the Transport Rules - This will get broken out more
#Get-TransportRule | Export-CSV c:\impact\transportrules.csv
Get-TransportRule | Export-Excel C:\Impact\Exchange_O365.xlsx -Append -WorksheetName "Transport Rules"

###############################################################################################################################################
###############################################################################################################################################
#Get-ImmutableID from Active Directory
$reportoutput=@()
$users = Get-ADUser -Filter * -Properties *
$users | Foreach-Object {



   $user = $_
    $immutableid = [System.Convert]::ToBase64String($user.ObjectGUID.tobytearray())
    $userid = $user | select @{Name='Access Rights';Expression={[string]::join(', ', $immutableid)}}



   $report = New-Object -TypeName PSObject
    $report | Add-Member -MemberType NoteProperty -Name 'UserPrincipalName' -Value $user.UserPrincipalName
    $report | Add-Member -MemberType NoteProperty -Name 'SamAccountName' -Value $user.samaccountname
    $report | Add-Member -MemberType NoteProperty -Name 'EmailAddress' -Value $user.emailaddress
    $report | Add-Member -MemberType NoteProperty -Name 'ImmutableID' -Value $immutableid
    $reportoutput += $report
}
 # Report
$reportoutput | Export-Excel C:\Impact\Exchange_O365.xlsx -Append -WorksheetName "ImmutableID"

#Perform a speedtest 

#############################################################################################################################
#############################################################################################################################
#GetAllMailboxRules
function ConvertTo-String ( $Object ) {
    if ( $Object -is 'Microsoft.Exchange.Data.Storage.Management.ADRecipientOrAddress[]' ) { return $Object -join ";" }
    else { return $Object }
}

function SanitizeString ( $string ) {
    if ( $null -eq $string ) { return $null, $null }
    else {
        $matches = [regex]::Matches( $string, '".*?"|\[.*?\]' )
        $name = $matches[0].Value
        $name = $name -replace '^"|"$', ''
        $email = $matches[1].Value
        if ( $email -match '\[SMTP:([^\]]+)\]' ) { $email = $Matches[1] }
        else {
            $exchangeUser = Get-User $name -ErrorAction SilentlyContinue
            if ( $exchangeUser ) { $email = $exchangeUser[0].WindowsEmailAddress }
            else { $email = $null }
        }
    return $name, $email
    }
}

$Mailboxes = Get-Mailbox -ResultSize Unlimited | Select Name, PrimarySmtpAddress, Guid

$Output = Foreach ( $Mailbox in $Mailboxes ) {
    $Rules = Get-InboxRule -Mailbox $Mailbox.Guid.Guid |
        Select MailboxOwnerID,Name,Enabled,Description,
        @{ n = 'From'; e = { ConvertTo-String( $_.From ) }},
        @{ n = 'RedirectTo'; e = { ConvertTo-String( $_.RedirectTo ) }},
        @{ n = 'ForwardTo'; e = { ConvertTo-String( $_.ForwardTo ) }} -ErrorAction SilentlyContinue

    Foreach ( $Rule in $Rules ) {
        $ruleFromName, $ruleFromEmail = SanitizeString ( $Rule.From )
        $ruleRedirectToName, $ruleRedirectToEmail = SanitizeString ( $Rule.RedirectTo )
        $ruleForwardToName, $ruleForwardToEmail = SanitizeString ( $Rule.ForwardTo )

        [PSCustomObject]@{
            MailboxOwnerID = $Rule.MailboxOwnerID
            Name = $Mailbox.Name
            EmailAddress = $Mailbox.PrimarySmtpAddress

            RuleName = $Rule.Name
            RuleEnabled = $Rule.Enabled

            FromAddress = $ruleFromEmail
            FromDisplayName = $ruleFromName
 
            ActivationDescription = $Rule.Description.ActivationDescription
            ExpiryDescription = $Rule.Description.ExpiryDescription

            RuleDescriptionTakeActions = $Rule.Description.RuleDescriptionTakeActions
            ActionDescriptions = $Rule.Description.ActionDescriptions -Join "; "

            RuleDescriptionIf = $Rule.Description.RuleDescriptionIf
            ConditionDescriptions = $Rule.Description.ConditionDescriptions -Join "; "

            RuleDescriptionExceptIf = $Rule.Description.RuleDescriptionIf
            ExceptionDescriptions = $Rule.Description.ExceptionDescriptions -Join "; "

            RuleDescriptionActivation = $Rule.Description.RuleDescriptionActivation
            RuleDescriptionExpiry = $Rule.Description.RuleDescriptionExpiry

            RedirectToAddress = $ruleRedirectToEmail
            RedirectToDisplayName = $ruleRedirectToName
            #RedirectToRoutingType = $Rule.RedirectTo.RoutingType

            ForwardToAddress = $ruleForwardToEmail
            ForwardToDisplayName = $ruleForwardToName
            #ForwardToRoutingType = $Rule.ForwardTo.RoutingType
        }
    }
}
$Output | Export-Excel C:\Impact\Exchange_O365.xlsx -Append -WorksheetName "AllMailboxRules"

##################################################################################################################################