function Backup-DFS
{
  <#
    Backs up .dfsutil backup .XML files of a DFS root

  #>
  [CmdletBinding()]
  param
  (
    [Parameter(Mandatory=$false, Position=0)]
    [System.String]
    $StorageLocation = 'c:\scripts\PowerShell\dfsbackup\'
  )
  
  $configurationContainer = ([adsi] 'LDAP://RootDSE').Get('ConfigurationNamingContext')
  $partitions = ([adsi] "LDAP://CN=Partitions,$configurationContainer").psbase.children
  [String]$Runtime = Get-Date -Format _yyyy.MMM.dd.HH.mmtt
  foreach($partition in $partitions)
  {
    if($partition.netbiosName -ne '')
    {
      $partitionDN=$partition.ncName
      $dnsName=$partitionDN.toString().replace('DC=','.').replace(',','').substring(1)
      $domain=$partition.netbiosName
      "`n$domain"
      New-Item $StorageLocation$domain$Runtime -ItemType Directory -ErrorAction SilentlyContinue
      $dfsContainer=[adsi] "LDAP://cn=Dfs-Configuration,cn=System,$partitionDN"
      $dfsRoots = $dfsContainer.psbase.Children
      foreach($dfsRoot in $dfsRoots)
      {
        $root=$dfsRoot.cn
        "`n$root"
        dfsutil root export "\\$dnsName\$root" "$StorageLocation$domain$Runtime\$root.xml"
      }
    }
  }
}