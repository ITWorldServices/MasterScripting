##
##
## This script will allow external emails to flow to a mail-enabled Public Folder.
## By default, Public Folders do not allow external emails. The below script will disbale the policy rule to allow external email and then give access to Anonymous for mail to be delivered
##
## Please modify the script below for the correct Public Folder location
#Connect to Exchnage Online

Connect-ExchangeOnline

## Disable the policy
## EXAMPLE: Set-MailPublicFolder -Identity “\Fax\Receipts” -EmailAddressPolicyEnabled $False
Set-MailPublicFolder -Identity “{PUBLIC FOLDER FULL PATH HERE}” -EmailAddressPolicyEnabled $False

## Enable Anonymous Access
## EXAMPLE: Add-PublicFolderClientPermission “\Fax\Receipts” -AccessRights CreateItems -User Anonymous
Add-PublicFolderClientPermission “{PUBLIC FOLDER FULL PATH HERE}” -AccessRights CreateItems -User Anonymous