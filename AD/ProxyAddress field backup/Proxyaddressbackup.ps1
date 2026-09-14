Import-Module ActiveDirectory
"DN;proxyAddress" | Out-File ".\proxyAddressesBackup.txt"
$Objects = Get-ADObject -LDAPFilter "(proxyAddresses=*)" -Properties proxyAddresses
ForEach ($Object In $Objects) {
  ForEach ($proxyAddress in $Object.proxyAddresses) {
    $Output = $Object.distinguishedName + ";" + $proxyAddress
    Write-Host $Output
    $Output | Out-File ".\proxyAddressesBackup.txt" -Append
  }
}