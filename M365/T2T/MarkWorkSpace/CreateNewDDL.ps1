Connect-MgGraph -Scopes "User.Read.All", "Group.ReadWrite.All"
Connect-ExchangeOnline
New-DynamicDistributionGroup -Name "Staff" -RecipientFilter {RecipientTypeDetails -eq "UserMailbox"}
