# --- Define export preference variables

$IncludeMailboxAccess = $true
$IncludeSendAs = $true
$IncludeSendOnBehalf = $true
$IncludeFolderDelegates = $true
$IncludeCommonFoldersOnly = $true
$DelegatesToSkip = "NT AUTHORITY\SELF","DOMAIN\BESADMIN","DOMAIN\Administrators"
$ExpandSecurityGroups = $false
$ExpandDistributionGroups = $false
$IncludeEntireForest = $true
$UseImportFile = $false
$ImportFile = "c:\scripts\UserList.csv"

# --- Initialize script functions

Function Write-Log ($LogString) {
    $LogStatus = $LogString.Split(":")[0]
    If ($LogStatus -eq "SUCCESS") {
        Write-Host $LogString -ForegroundColor Green
        $LogString | Out-File $RunLog -Append  }
    If ($LogStatus -eq "INFO") {
        Write-Host "$LogString" -ForegroundColor Cyan
        $LogString | Out-File $RunLog -Append }
    If ($LogStatus -eq "ALERT") {
        Write-Host $LogString -ForegroundColor Yellow
        $LogString | Out-File $RunLog -Append }
    If ($LogStatus -eq "ERROR") {
        Write-Host $LogString -BackgroundColor Red
        $LogString | Out-File $RunLog -Append
        "`n" | Out-File $ErrorLog -Append
        $LogString | Out-File $ErrorLog -Append }
    If ($LogStatus -eq "AUDIT") {
        Write-Host $LogString -ForegroundColor DarkGray
        $LogString | Out-File $RunLog -Append  }
    If ($LogStatus -eq "") {
        Write-Host ""
        Write-Output "`n" | Out-File $RunLog -Append }
}

Function Check-Delegates ([string]$DelegateID, $ExportFile) {
    $CheckDelegate = Get-Recipient $DelegateID -ErrorAction SilentlyContinue
    If ($CheckDelegate -eq $null) {
        $CheckDelegate = Get-Group $DelegateID -ErrorAction SilentlyContinue }
    If ($CheckDelegate -ne $null) {
        If (($CheckDelegate.RecipientType -like "Mail*" -and $ExpandDistributionGroups -eq $false) -or $CheckDelegate.RecipientType -like "*Mailbox") {
            $DelegateName = $CheckDelegate.Name
            $DelegateEmail = $CheckDelegate.PrimarySmtpAddress
            "$MailboxName,$MailboxEmail,$DelegateName,$DelegateEmail,$DelegateAccess" | Out-File $ExportFile -Append }
        If ($CheckDelegate.RecipientType -like "Mail*" -and $CheckDelegate.RecipientType -like "*Group" -and $ExpandDistributionGroups -eq $true) {
            Write-Log "ALERT: Expand distribution group membership. [$($CheckDelegate.Name)]"
            ForEach ($Member in Get-DistributionGroupMember $CheckDelegate.Name -ResultSize Unlimited) {
                $CheckMember = Get-Recipient $Member -ErrorAction SilentlyContinue
                If ($CheckMember -ne $null) {
                    $DelegateName = $DelegateID + ":" + $CheckMember.Name
                    $DelegateEmail = $CheckMember.PrimarySmtpAddress
                    "$MailboxName,$MailboxEmail,$DelegateName,$DelegateEmail,$DelegateAccess" | Out-File $ExportFile -Append } } }
        If ($CheckDelegate.RecipientType -eq "Group" -and $ExpandSecurityGroups -eq $true) {
            Write-Log "ALERT: Expand security group membership. [$($CheckDelegate.Name)]"
            ForEach ($Member in (Get-Group $DelegateID).Members) {
                $CheckMember = Get-Recipient $Member -ErrorAction SilentlyContinue
                If ($CheckMember -ne $null) {
                    $DelegateName = $DelegateID + ":" + $CheckMember.Name
                    $DelegateEmail = $CheckMember.PrimarySmtpAddress
                    "$MailboxName,$MailboxEmail,$DelegateName,$DelegateEmail,$DelegateAccess" | Out-File $ExportFile -Append } } } }      
 }

<################################################################################
SCRIPT BEGIN
################################################################################>
cls; Write-Host "$ScriptInfo`n" -ForegroundColor White

# --- Initialize log files

$TimeStamp = Get-Date -Format yyMMddhhmmss
$RunLog = $TimeStamp + "_Export-MailboxPermissions_RunLog.txt"
$ErrorLog = $TimeStamp + "_Export-MailboxPermissions_ErrorLog.txt"

# --- Initialize environment

$EmsVersion = Get-PSSnapin -Registered | Where { $_.Name -like "*Exchange*" }
If ($EmsVersion -like "*Powershell.E2010*") { $EmsVersion = "EX2010" }
If ($EmsVersion -like "*Powershell.Admin*") { $EmsVersion = "EX2007" }
If ($IncludeFolderDelegates -eq $true -and $EmsVersion -eq "EX2007") {
    Write-Log "WARNING: Folder delegate permissions cannot be retrieved from Exchange 2007 and will not be exported.`n
               Consider using PFDAVAdmin to export folder permissions from Exchange 2007.`n"
    $IncludeFolderDelegates = $false }
If ($IncludeEntireForest -eq $true -and $EmsVersion -eq "EX2010") { Set-AdServerSettings -ViewEntireForest $true }
If ($IncludeEntireForest -eq $true -and $EmsVersion -eq "EX2007") { $AdminSessionADSettings.ViewEntireForest = $true } 

# --- Initialize export files

$MailboxAccessExport = ".\MailboxAccess.csv"
$SendAsExport = ".\MailboxSendAs.csv"
$SendOnBehalfExport = ".\MailboxSendOnBehalf.csv"
$FolderDelegateExport = ".\MailboxFolderDelegate.csv"
If ($IncludeMailboxAccess -eq $true) { "MailboxName,MailboxEmail,DelegateName,DelegateEmail,DelegateAccess" | Out-File $MailboxAccessExport }
If ($IncludeSendAs -eq $true) { "MailboxName,MailboxEmail,DelegateName,DelegateEmail,DelegateAccess" | Out-File $SendAsExport }
If ($IncludeSendOnBehalf -eq $true) { "MailboxName,MailboxEmail,DelegateName,DelegateEmail,DelegateAccess" | Out-File $SendOnBehalfExport }
If ($IncludeFolderDelegates -eq $true) { "MailboxName,MailboxEmail,FolderLocation,DelegateName,DelegateEmail,DelegateAccess" | Out-File $FolderDelegateExport }
Get-Date | Out-File $ErrorLog

# --- Retrieve mailboxes

If ($UseImportFile -eq $true) {
    Write-Log "INFO: Retrieve mailboxes from user list. [$ImportFile]"
    $Mailboxes = (Import-Csv $ImportFile | Select PrimarySmtpAddress) }
Else {
    Write-Log "INFO: Retrieving mailboxes from Exchange..."
    $Mailboxes = (Get-Mailbox -ResultSize Unlimited | Select PrimarySmtpAddress) }
$MailboxCount = $Mailboxes.Count; $Progress = 0
Write-Log "SUCCESS: Found $MailboxCount Mailboxes."

# --- Process mailboxes

ForEach ($Mailbox in $Mailboxes) {
    [string]$MailboxEmail = $Mailbox.PrimarySmtpAddress
    $CheckMailbox = Get-Recipient $MailboxEmail -ErrorAction SilentlyContinue
    If ($CheckMailbox -eq $null) { Write-Log "ERROR: Mailbox not found. [$MailboxEmail]"; Continue }
    [string]$MailboxName = $CheckMailbox.Name
    [string]$MailboxDN = $CheckMailbox.DistinguishedName
    $Progress = $Progress + 1
    Write-Log ""; Write-Log "INFO: Audit mailbox $Progress of $MailboxCount. [$MailboxEmail]"

    # --- Export mailbox access permissions

    If ($IncludeMailboxAccess -eq $true) {
        Write-Log "AUDIT: Mailbox access permissions..."
        $Delegates = @()
        $Delegates = (Get-MailboxPermission $MailboxDN | Where { $DelegatesToSkip -notcontains $_.User -and $_.IsInherited -eq $false })
        If ($Delegates -ne $null) {
            ForEach ($Delegate in $Delegates) {
                $DelegateAccess = $Delegate.AccessRights
                Check-Delegates $Delegate.User $MailboxAccessExport } } }

    # --- Export SendAs permissions

    If ($IncludeSendAs -eq $true) {
        Write-Log "AUDIT: Send As permissions..."
        $Delegates = @()
        $Delegates = Get-ADPermission $MailboxDN | Where { $DelegatesToSkip -notcontains $_.User -and $_.ExtendedRights -like "*send-as*" }
        If ($Delegates -ne $null) {
            ForEach ($Delegate in $Delegates) {
                $DelegateAccess = "SendAs" 
                Check-Delegates $Delegate.User $SendAsExport } } }

    # --- Export SendOnBehalf permissions

    If ($IncludeSendOnBehalf -eq $true) {
        Write-Log "AUDIT: Send On Behalf permissions..."
        $Delegates = @()
        $Delegates = (Get-Mailbox $MailboxDN).GrantSendOnBehalfTo
        If ($Delegates -ne $null) {
            ForEach ($Delegate in $Delegates) {
                $DelegateAccess = "SendOnBehalf"
                Check-Delegates $Delegate.Name $SendOnBehalfExport } } }

    # --- Export folder permissions

    If ($IncludeFolderDelegates -eq $true) {
        Write-Log "AUDIT: Folder delegate permissions."
        If ($IncludeCommonFoldersOnly -eq $true) { $Folders = Get-MailboxFolderStatistics $MailboxDN | Where { $_.FolderPath -eq "/Top of Information Store" -or $_.FolderPath -eq "/Inbox" -or $_.FolderPath -eq "/Calendar" } }
        Else { $Folders = Get-MailboxFolderStatistics $MailboxDN }
        ForEach ($Folder in $Folders) {
            $FolderPath = $Folder.FolderPath.Replace("/","\")
            If ($FolderPath -eq "\Top of Information Store") { $FolderPath = "\" }
            $FolderLocation = $MailboxEmail + ":" + $FolderPath
            $FolderPermissions = Get-MailboxFolderPermission $FolderLocation -ErrorAction SilentlyContinue
            If ($FolderPermissions -ne $null) {
                ForEach ($Permission in $FolderPermissions) {
                    [string]$FolderDelegate = $Permission.User
                    If ($FolderDelegate -ne "Default" -and $FolderDelegate -ne "Anonymous") {
                        $CheckDelegate = Get-Recipient $FolderDelegate -ErrorAction SilentlyContinue
                        If ($CheckDelegate -ne $null) {
                            $DelegateName = $CheckDelegate.Name
                            $DelegateEmail = $CheckDelegate.PrimarySmtpAddress
                            $DelegateAccess = $Permission.AccessRights
                            If ($DelegateName -ne $MailboxName) { "$MailboxName,$MailboxEmail,$FolderPath,$DelegateName,$DelegateEmail,$DelegateAccess" | Out-File $FolderDelegateExport -Append } } } } } } }
}

<###############################################################################
SCRIPT END
###############################################################################>

Write-Log ""; Write-Log "SUCCESS: Permission export complete!"