import-module ActiveDirectory

$CSVFile = "AllMailboxPermissions.csv"

'DisplayName' + ',' + 'Email Address' + ',' + 'Full Access' + ',' + 'Send As' + ',' + 'Send On Behalf Of' | out-file -FilePath $CSVfile -Force
 
$Mailboxes = Get-Mailbox -resultsize unlimited | Select Identity,Alias,DisplayName,DistinguishedName,WindowsEmailAddress

ForEach ( $Mailbox in $Mailboxes ) {

    $SendOnBehalfOf = Get-mailbox $Mailbox.identity | % { $_.GrantSendOnBehalfTo }
    
    $SendAs = Get-ADPermission $Mailbox.identity |
        where {( $_.ExtendedRights -like "*Send*" )
        -and -not ( $_.User -like "NT AUTHORITY\SELF" )
        -and -not ( $_.User -like "S-1-5-21*" )} |
        % { $_.User }
    
    $FullAccess = Get-MailboxPermission $Mailbox.Identity |
        ?{( $_.IsInherited -eq $False )
        -and -not ( $_.User -match "NT AUTHORITY" )} |
        Select User,Identity,@{ Name="AccessRights"; Expression={ $_.AccessRights }} |
        % { $_.User }
    
    $Mailbox.DisplayName + ',' + $Mailbox.WindowsEmailAddress + ',' + $FullAccess + ',' + $SendAs + ',' + $SendOnBehalfOf | out-file -FilePath $CSVfile -Append
}