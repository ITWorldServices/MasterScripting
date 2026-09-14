#Install the Excel Plug-in. We want to be able to export to Excel.
# Check if ImportExcel module is installed
#$importExcelModuleInstalled = Get-Module -Name ImportExcel -ListAvailable

#if (!$importExcelModuleInstalled) {
#    Write-Host "ImportExcel module is not installed. Installing..."
#    Install-Module -Name ImportExcel -Scope CurrentUser -Force
#} else {
#    Write-Host "ImportExcel module is already installed."
#}
#Install-Module -Name ImportExcel -Scope CurrentUser -Force

# Import the required modules
Import-Module ActiveDirectory
#Import-Module ImportExcel

#Get-ImmutableID from Active Directory
$reportoutput=@()
$users = Get-ADUser -Filter * -Properties *
$users | Foreach-Object {



   $user = $_
    $immutableid = [System.Convert]::ToBase64String($user.ObjectGUID.tobytearray())
    $userid = $user | select @{Name='Access Rights';Expression={[string]::join(', ', $immutableid)}}



   $report = New-Object -TypeName PSObject
    $report | Add-Member -MemberType NoteProperty -Name 'UserPrincipalName' -Value $user.UserPrincipalName
    $report | Add-Member -MemberType NoteProperty -Name 'SamAccountName' -Value $user.samaccountname
    $report | Add-Member -MemberType NoteProperty -Name 'EmailAddress' -Value $user.emailaddress
    $report | Add-Member -MemberType NoteProperty -Name 'ImmutableID' -Value $immutableid
    $reportoutput += $report
}
 # Report
$SyncFileName = "C:\temp\Syncusers.csv"
if (Test-Path $SyncFileName) {
   Remove-Item $SyncFileName -verbose
}
#$reportoutput | Export-Excel C:\temp\Exchange_O365_$(get-date -f yyyy-MM-dd).xlsx -Append -WorksheetName "ImmutableID"

$reportoutput | Export-csv $SyncFileName