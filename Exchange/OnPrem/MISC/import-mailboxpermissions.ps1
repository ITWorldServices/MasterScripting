# --- Define import preference variables

$IncludeMailboxAccess = $true
$IncludeSendAs = $true
$IncludeSendOnBehalf = $false
$IncludeFolderDelegates = $false

<################################################################################
SCRIPT BEGIN
################################################################################>
cls; Write-Host "$ScriptInfo`n" -ForegroundColor White

# --- Import mailbox access permissions

If ($IncludeMailboxAccess -eq $true) {
    Write-Host "Applying mailbox access permissions..."
    Import-Csv ".\MailboxAccess.csv" | ForEach {
        Add-MailboxPermission $_.MailboxEmail -User $_.DelegateEmail -AccessRights $_.DelegateAccess } }

# --- Import SendAs permissions

If ($IncludeSendAs -eq $true) {
    Write-Host "Applying SendAs permissions..."
    Import-Csv ".\MailboxSendAs.csv" | ForEach {
        Add-RecipientPermission $_.MailboxEmail -Trustee $_.DelegateEmail -AccessRights "SendAs" -Confirm:$false } }

# --- Import SendOnBehalf permissions

If ($IncludeSendOnBehalf -eq $true) {
    Write-Host "Applying SendOnBehalf permissions..."
    Import-Csv ".\MailboxSendOnBehalf.csv" | ForEach {
        Set-Mailbox $_.MailboxEmail -GrantSendOnBehalfTo @{Add=$_.DelegateEmail} } }

# --- Import folder delegate permissions

If ($IncludeFolderDelegates -eq $true) {
    Write-Host "Applying folder delegate permissions..." -ForegroundColor Yellow 
    Import-Csv ".\MailboxFolderDelegate.csv" | ForEach {
        $FolderPath = $_.MailboxEmail + ":" + $_.FolderLocation
        Add-MailboxFolderPermission $FolderPath -User $_.DelegateEmail -AccessRights $_.DelegateAccess } }

<###############################################################################
SCRIPT END
###############################################################################>

Write-Host "Done!" -ForegroundColor Green