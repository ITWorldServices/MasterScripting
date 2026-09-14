$Result=@()
#Get all office 365 personal sites
$oneDriveSites = Get-SPOSite -IncludePersonalSite $true -Limit All -Filter "Url -like '-my.sharepoint.com/personal/'"
$oneDriveSites | ForEach-Object {
$site = $_
$Result += New-Object PSObject -property @{ 
UserName = $site.Owner
OneDriveSiteUrl = $site.URL
}
}
$Result | Select UserName, OneDriveSiteUrl |
Export-CSV "LOCALPATHINSERTEDHERE\VoM-OneDrive-for-Business-Users.csv" -NoTypeInformation -Encoding UTF8
