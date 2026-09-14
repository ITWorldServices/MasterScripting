Import-CSV "MASTER.csv" | % {Set-ADAccountPassword -Identity $_.SamAccountName -Reset -NewPassword (ConvertTo-SecureString -AsPlainText $_.NewPassword -Force)}
