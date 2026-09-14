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
$Output | Export-CSV TestInboxRules.csv -NoType
