Import-Csv proxyAddresses-test.csv | ForEach-Object{
    $upn=$_.UserPrincipalName;
    $aduser=Get-ADUser -Filter { UserPrincipalName -Eq $upn }
    $proxy=$_.ProxyAddresses -split ","
    $proxy | ForEach-Object{
        $user=[ADSI]"LDAP://$($aduser.distinguishedname)"
        $ads_property_append=3
        $user.Putex($ADS_PROPERTY_APPEND,"proxyaddresses",@($proxy))
        $user.setinfo()
    }
}
#Technical Project Team Is the Best.