$userid='sending_user@sending_domain.com'
$creds=Get-Credential $userid
import-csv 'All_Active_Users_With_Attributes.csv' | % {
$mobile = $_.mobilephone
$body = 'Hello '+$($_.name)+','+[Environment]::NewLine+[Environment]::NewLine+'First Hospitality Group has enabled a Self Service Password Reset Portal where you can reset your FHG password using your phone and a password reset code.'+[Environment]::NewLine+[Environment]::NewLine+'You may navigate to https://ids1075.nerdio.net:5000/ (or http://bit.ly/FHGPass for a shorter version) from any computer (except for Nerdio Desktops) to reset your FHG password.'+[Environment]::NewLine+[Environment]::NewLine+'You will need the following:'+[Environment]::NewLine+'     Username: '+$($_.samaccountname)+[Environment]::NewLine+'     Passcode: '+$($_.ipphone)+[Environment]::NewLine+'     Mobile Phone: '+$($_.mobilephone)+[Environment]::NewLine+[Environment]::NewLine+'You should receive a separate email with step-by-step instructions on how to use the Self Service Password Reset Portal.'
if ($mobile.Length -gt 1) {
    Write-Output "Sent Email to: "+$_.emailaddress
    send-mailmessage -To $_.emailaddress -Subject 'Self Service Password Reset Passcode' -Body $body -UseSsl -Port 25 -SmtpServer 'fhginc-com.mail.protection.outlook.com' -Priority High -From $userid -Credential $creds
    }
}
