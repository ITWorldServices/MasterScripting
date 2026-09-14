Export Mailbox Permissions 
Export-MailboxPermissions.ps1 should be run from on-premises Exchange Management Shell (EMS) and will collect mailbox access, send as, send on behalf, and folder delegate permissions into separate CSV datasets.  Enumerating folder delegates can take considerable time, so auditing only common folders (Inbox, Calendar) is enabled by default. Mailbox permissions in O365 can only be assigned using mail-enabled objects, so script offers options for expanding group memberships.

Export Preferences
The following script variables can be modified according to requirements or preferences:<br/>

Variable	Description	Values<br/>
$IncludeMailboxAccess<br/>
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Specifies whether on-premises mailbox access permissions (e.g. FullAccess, ReadPermission, etc.) are audited and exported for migration.<br/>
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;$true (Default)<br/>
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;$false<br/>

$IncludeSendAs<br/>
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Specifies whether on-premises "Send As" permissions are audited and exported for migration.<br/>
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;$true (Default)<br/>
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;$false<br/>
  
$IncludeSendOnBehalf<br/>
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Specifies whether on-premises "Send On Behalf" permissions are audited and exported for migration.<br/>
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;$true (Default)<br/>
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;$false<br/>

$IncludeFolderDelegates<br/>
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Specifies whether on-premises Outlook folder delegates are audited and exported for migration.  Enumerating delegate permissions for all Outlook folders can take considerable time in large environments.<br/>
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;$true (Default)<br/>
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;$false<br/>

$IncludeCommonFoldersOnly<br/>
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;If preference variable $IncludeFolderDelegates is $true, specifies whether only Inbox and Calendar delegate permissions are exported for migration.<br/>
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;$true (Default)<br/>
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;$false<br/>
  
$DelegatesToSkip<br/>
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Specifies service accounts which should not be exported for migration.	See script for examples.<br/>

$ExpandDistributionGroups<br/>
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Specifies whether mail-enabled groups are expanded so their members can be explicitly migrated. Mail-enabled groups are synchronized to O365 and permissions should be retained without the need for expansion.<br/>
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;$true<br/>
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;$false (Default)<br/>

$ExpandSecurityGroups<br/>
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Specifies whether security groups are expanded so their members can be explicitly migrated.  Security groups are synchronized to O365, but cannot be used to define mailbox permissions.  If possible, consider mail-enabling security groups and hiding them from GAL.	<br/>
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;$true<br/>
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;$false (Default)<br/>

$IncludeEntireForest<br/>
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Specifies whether mailbox permissions are audited in current AD domain or entire Exchange forest.<br/>	
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;$true (Default)<br/>
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;$false<br/>

$UseImportFile<br/>
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Specifies whether mailbox permissions are audited for a list of users.  Import file requires "PrimarySmtpAddress" column.<br/>
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;$true<br/>
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;$false (Default)<br/>

$ImportFile<br/>
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Specifies the path and file name for the user import CSV.	c:\path\file.csv<br/>

Import Mailbox Permissions<br/>
Import-MailboxPermissions.ps1 should be run from O365 remote PowerShell after mailboxes have been provisioned and will re-apply permissions according to collected on-premises datasets.  Import of each export file can be toggled "$true" or "$false" in the script and re-applied separately if needed. 
