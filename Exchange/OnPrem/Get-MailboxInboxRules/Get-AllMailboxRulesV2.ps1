<# K and BK Script to Pull Exchange On-Premise Mailbox Rules v1.0 #>

$Mailboxes = Get-Mailbox -ResultSize Unlimited  

$Output = Foreach ($Mailbox in $Mailboxes) {
  $Rules = Get-InboxRule -Mailbox $Mailbox.UserPrincipalName  #| Select -Property MailboxOwnerID,Name,Enabled,From,Description,RedirectTo,ForwardTo
  Foreach ($Rule in $Rules) {
    [PSCustomObject]@{

      MailboxOwnerID = $Rule.MailboxOwnerID
      Name = $Mailbox.Name
      EmailAddress = $Mailbox.EmailAddress
      
      RuleName = $Rule.Name
      RuleEnabled = $Rule.Enabled
      
      FromAddress = $Rule.From | % Address  -Join "; "
      FromDisplayName = $Rule.From.DisplayName -Join "; "
      
      ActivationDescription = $Rule.Description.ActivationDescription   # -Join "; "
      ExpiryDescription = $Rule.Description.ExpiryDescription
      
      RuleDescriptionTakeActions = $Rule.Description.RuleDescriptionTakeActions
      ActionDescriptions = $Rule.Description.ActionDescriptions -Join "; "
      
      RuleDescriptionIf = $Rule.Description.RuleDescriptionIf
      ConditionDescriptions = $Rule.Description.ConditionDescriptions -Join "; "
      
      RuleDescriptionExceptIf = $Rule.Description.RuleDescriptionIf
      ExceptionDescriptions = $Rule.Description.ExceptionDescriptions -Join "; "
      
      RuleDescriptionActivation = $Rule.Description.RuleDescriptionActivation
      RuleDescriptionExpiry = $Rule.Description.RuleDescriptionExpiry
      
      RedirectToAddress	= $Rule.RedirectTo | % Address -Join "; "
      RedirectToDisplayName = $Rule.RedirectTo.DisplayName
      RedirectToRoutingType = $Rule.RedirectTo.RoutingType
      
      ForwardToAddress = $Rule.ForwardTo | % Address -Join "; "
      ForwardToDisplayName = $Rule.ForwardTo.DisplayName
      ForwardToRoutingType = $Rule.ForwardTo.RoutingType
    }
  }
}
$Output | Export-CSV C:\Impact\NAMEYOURFILES.csv -NoType
