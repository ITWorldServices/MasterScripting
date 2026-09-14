$comps = get-adcomputer -filter {(OperatingSystem -like "*windows*server*") -and (Enabled -eq "True")} -Properties *
"You have [{0}] servers in domain [{1}]" -f $comps.count, (get-addomain).dnsroot

$today = get-date
$monthago = $today.AddDays(-30)
"Looking for systems that have not logged in since $monthago"

foreach ($comp in $comps) 
{
   if ($comp.lastlogondate -lt $monthago)
      {"Computer [$comp] suspect" 
       "last logon $($comp.lastlogon)"
       ""}
}
