$credential = Get-Credential  unitedalerts@chartertsp.com

## Define the Send-MailMessage parameters
$mail = @{
    SmtpServer                 = 'smtp.office365.com'
    Port                       = '587'
    UseSSL                     = $true
    Credential                 = $credential
    From                       = 'unitedalerts@chartertsp.com'
    To                         = 'unitedalerts@chartertsp.com'
    Subject                    = "SMTP Client Submission - $(Get-Date -Format g)"
    Body                       = 'This is a test email using SMTP Client Submission'
    DeliveryNotificationOption = 'OnFailure', 'OnSuccess'
}

Send-MailMessage @mail