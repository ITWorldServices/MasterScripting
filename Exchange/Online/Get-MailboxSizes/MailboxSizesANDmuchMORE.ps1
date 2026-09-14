<#
    .SYNOPSIS
    .\MailboxSizeReport365.ps1

    .DESCRIPTION
    Connect to Exchange Online PowerShell first.
    The script exports a Mailbox Size Report for all Microsoft 365 mailboxes
    to a CSV file. You can also export a single mailbox or WildCard as an option.

    .LINK
    o365info.com/microsoft-365-mailbox-size-report

    .NOTES
    Written By: o365info
    Website:    o365info.com

    .CHANGELOG
    V1.00, 03/24/2023 - Initial version
#>

Write-host "

Mailbox Size Report 365
----------------------------

1.Export to CSV File (OFFICE 365)

2.Enter the Mailbox Name with Wild Card (Export) (OFFICE 365)"-ForeGround "Cyan"

#----------------
# Script
#----------------

Write-Host "               "

$number = Read-Host "Choose The Task"
$output = @()
switch ($number) {

    1 {

        $i = 0 
        $CSVfile = Read-Host "Enter the Path of CSV file (Eg. C:\Report.csv)" 

        $AllMailbox = Get-mailbox -Resultsize Unlimited

        Foreach ($Mbx in $AllMailbox) {

            $Stats = Get-mailboxStatistics -Identity $Mbx.distinguishedname -WarningAction SilentlyContinue

            if ($Mbx.ArchiveName.count -eq "0") {
                $ArchiveTotalItemSize = $null
                $ArchiveTotalItemCount = $null
            }
            if ($Mbx.ArchiveName -ge "1") {
                $MbxArchiveStats = Get-mailboxstatistics $Mbx.distinguishedname -Archive -WarningAction SilentlyContinue
                $ArchiveTotalItemSize = $MbxArchiveStats.TotalItemSize
                $ArchiveTotalItemCount = $MbxArchiveStats.BigFunnelMessageCount
            }

            $userObj = New-Object PSObject

            $userObj | Add-Member NoteProperty -Name "Display Name" -Value $mbx.displayname
            $userObj | Add-Member NoteProperty -Name "Alias" -Value $Mbx.Alias
            $userObj | Add-Member NoteProperty -Name "SamAccountName" -Value $Mbx.SamAccountName
            $userObj | Add-Member NoteProperty -Name "RecipientType" -Value $Mbx.RecipientTypeDetails
            $userObj | Add-Member NoteProperty -Name "Recipient OU" -Value $Mbx.OrganizationalUnit
            $userObj | Add-Member NoteProperty -Name "Primary SMTP address" -Value $Mbx.PrimarySmtpAddress
            $userObj | Add-Member NoteProperty -Name "Email Addresses" -Value ($Mbx.EmailAddresses -join ",")
            $userObj | Add-Member NoteProperty -Name "Database" -Value $Stats.Database
            $userObj | Add-Member NoteProperty -Name "ServerName" -Value $Stats.ServerName
            $userObj | Add-Member NoteProperty -Name "TotalItemSize" -Value $Stats.TotalItemSize
            $userObj | Add-Member NoteProperty -Name "ItemCount" -Value $Stats.ItemCount
            $userObj | Add-Member NoteProperty -Name "DeletedItemCount" -Value $Stats.DeletedItemCount
            $userObj | Add-Member NoteProperty -Name "TotalDeletedItemSize" -Value $Stats.TotalDeletedItemSize
            $userObj | Add-Member NoteProperty -Name "ProhibitSendReceiveQuota-In-MB" -Value $Mbx.ProhibitSendReceiveQuota
            $userObj | Add-Member NoteProperty -Name "UseDatabaseQuotaDefaults" -Value $Mbx.UseDatabaseQuotaDefaults
            $userObj | Add-Member NoteProperty -Name "LastLogonTime" -Value $Stats.LastLogonTime
            $userObj | Add-Member NoteProperty -Name "ArchiveName" -Value ($Mbx.ArchiveName -join ",")
            $userObj | Add-Member NoteProperty -Name "ArchiveStatus" -Value $Mbx.ArchiveStatus
            $userObj | Add-Member NoteProperty -Name "ArchiveState" -Value $Mbx.ArchiveState 
            $userObj | Add-Member NoteProperty -Name "ArchiveQuota" -Value $Mbx.ArchiveQuota
            $userObj | Add-Member NoteProperty -Name "ArchiveTotalItemSize" -Value $ArchiveTotalItemSize
            $userObj | Add-Member NoteProperty -Name "ArchiveTotalItemCount" -Value $ArchiveTotalItemCount

            $output += $UserObj  
            # Update Counters and Write Progress
            $i++
            if ($AllMailbox.Count -ge 1) {
                Write-Progress -Activity "Scanning Mailboxes . . ." -Status "Scanned: $i of $($AllMailbox.Count)" -PercentComplete ($i / $AllMailbox.Count * 100)
            }
        }

        $output | Export-csv -Path $CSVfile -NoTypeInformation -Encoding UTF8 #-Delimiter ","

        ; Break
    }

    2 {
        $i = 0
        $CSVfile = Read-Host "Enter the Path of CSV file (Eg. C:\DG.csv)" 

        $MailboxName = Read-Host "Enter the Mailbox name or Range (Eg. Mailboxname , Mi*,*Mik)"

        $AllMailbox = Get-mailbox $MailboxName -Resultsize Unlimited

        Foreach ($Mbx in $AllMailbox) {

            $Stats = Get-mailboxStatistics -Identity $Mbx.distinguishedname -WarningAction SilentlyContinue

            if ($Mbx.ArchiveName.count -eq "0") {
                $ArchiveTotalItemSize = $null
                $ArchiveTotalItemCount = $null
            }
            if ($Mbx.ArchiveName -ge "1") {
                $MbxArchiveStats = Get-mailboxstatistics $Mbx.distinguishedname -Archive -WarningAction SilentlyContinue
                $ArchiveTotalItemSize = $MbxArchiveStats.TotalItemSize
                $ArchiveTotalItemCount = $MbxArchiveStats.BigFunnelMessageCount
            }

            $userObj = New-Object PSObject

            $userObj | Add-Member NoteProperty -Name "Display Name" -Value $mbx.displayname
            $userObj | Add-Member NoteProperty -Name "Alias" -Value $Mbx.Alias
            $userObj | Add-Member NoteProperty -Name "SamAccountName" -Value $Mbx.SamAccountName
            $userObj | Add-Member NoteProperty -Name "RecipientType" -Value $Mbx.RecipientTypeDetails
            $userObj | Add-Member NoteProperty -Name "Recipient OU" -Value $Mbx.OrganizationalUnit
            $userObj | Add-Member NoteProperty -Name "Primary SMTP address" -Value $Mbx.PrimarySmtpAddress
            $userObj | Add-Member NoteProperty -Name "Email Addresses" -Value ($Mbx.EmailAddresses -join ",")
            $userObj | Add-Member NoteProperty -Name "Database" -Value $Stats.Database
            $userObj | Add-Member NoteProperty -Name "ServerName" -Value $Stats.ServerName
            $userObj | Add-Member NoteProperty -Name "TotalItemSize" -Value $Stats.TotalItemSize
            $userObj | Add-Member NoteProperty -Name "ItemCount" -Value $Stats.ItemCount
            $userObj | Add-Member NoteProperty -Name "DeletedItemCount" -Value $Stats.DeletedItemCount
            $userObj | Add-Member NoteProperty -Name "TotalDeletedItemSize" -Value $Stats.TotalDeletedItemSize
            $userObj | Add-Member NoteProperty -Name "ProhibitSendReceiveQuota-In-MB" -Value $Mbx.ProhibitSendReceiveQuota
            $userObj | Add-Member NoteProperty -Name "UseDatabaseQuotaDefaults" -Value $Mbx.UseDatabaseQuotaDefaults
            $userObj | Add-Member NoteProperty -Name "LastLogonTime" -Value $Stats.LastLogonTime
            $userObj | Add-Member NoteProperty -Name "ArchiveName" -Value ($Mbx.ArchiveName -join ",")
            $userObj | Add-Member NoteProperty -Name "ArchiveStatus" -Value $Mbx.ArchiveStatus
            $userObj | Add-Member NoteProperty -Name "ArchiveState" -Value $Mbx.ArchiveState 
            $userObj | Add-Member NoteProperty -Name "ArchiveQuota" -Value $Mbx.ArchiveQuota
            $userObj | Add-Member NoteProperty -Name "ArchiveTotalItemSize" -Value $ArchiveTotalItemSize
            $userObj | Add-Member NoteProperty -Name "ArchiveTotalItemCount" -Value $ArchiveTotalItemCount

            $output += $UserObj  
            # Update Counters and Write Progress
            $i++
            if ($AllMailbox.Count -ge 1) {
                Write-Progress -Activity "Scanning Mailboxes . . ." -Status "Scanned: $i of $($AllMailbox.Count)" -PercentComplete ($i / $AllMailbox.Count * 100) -ErrorAction SilentlyContinue
            }
        }

        $output | Export-csv -Path $CSVfile -NoTypeInformation -Encoding UTF8 #-Delimiter ","

        ; Break
    }

    Default { Write-Host "No matches found , Enter Options 1 or 2" -ForeGround "red" }

}