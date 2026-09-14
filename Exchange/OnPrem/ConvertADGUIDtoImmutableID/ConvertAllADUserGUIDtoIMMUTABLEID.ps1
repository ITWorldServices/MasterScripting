$reportoutput=@()
$users = Get-ADUser -Filter * -Properties *
$users | Foreach-Object {



   $user = $_
    $immutableid = [System.Convert]::ToBase64String($user.ObjectGUID.tobytearray())
    $userid = $user | select @{Name='Access Rights';Expression={[string]::join(', ', $immutableid)}}



   $report = New-Object -TypeName PSObject
    $report | Add-Member -MemberType NoteProperty -Name 'UserPrincipalName' -Value $user.UserPrincipalName
    $report | Add-Member -MemberType NoteProperty -Name 'SamAccountName' -Value $user.samaccountname
    $report | Add-Member -MemberType NoteProperty -Name 'ImmutableID' -Value $immutableid
    $reportoutput += $report
}
 # Report
$reportoutput | Export-Csv -Path c:\impact\ImmutableID4AD.csv -NoTypeInformation -Encoding UTF8