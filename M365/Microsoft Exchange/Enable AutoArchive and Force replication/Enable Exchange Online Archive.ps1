#The below steps will enable the AutoExpanding Archive feature on the tenant, allow that to be turned on for the users, and then the step to force a replication

#First connect to Exchange
Connect-ExchangeOnline

#Enable Autoexpanding Archive for the Organization. This is at an Organization level
Set-OrganizationConfig -AutoExpandingArchive

#Now enable Autoexpanding Archive for a user
enable-mailbox <email address> -AutoExpandingArchive

#Now force Exchange to start migrating the user. I run this about every 30 minutes.
#An error may show that the Archive is Offline. That happens. Wait 5 minutes and run the process again

Start-ManagedFolderAssistant -Identity <email address>