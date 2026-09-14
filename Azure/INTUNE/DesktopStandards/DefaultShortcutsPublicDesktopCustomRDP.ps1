# Add APP1.UC.GROUP shortcut
$new_object_1 = New-Object -ComObject WScript.Shell
$destination = $new_object_1.SpecialFolders.Item("AllUsersDesktop")
$source_path = Join-Path -Path $destination -ChildPath "\\App1.url"
$source = $new_object_1.CreateShortcut($source_path)
$source.TargetPath = "https://app1.uc.group/RDWeb/Pages/en-US/login.aspx?ReturnUrl=/RDWeb/Pages/en-US/Default.aspx"
$source.Save()

# Add APP2.UC.GROUP shortcut
$new_object_2 = New-Object -ComObject WScript.Shell
$destination = $new_object_2.SpecialFolders.Item("AllUsersDesktop")
$source_path = Join-Path -Path $destination -ChildPath "\\App2.url"
$source = $new_object_2.CreateShortcut($source_path)
$source.TargetPath = "https://app2.uc.group/RDWeb/Pages/en-US/login.aspx?ReturnUrl=/RDWeb/Pages/en-US/Default.aspx"
$source.Save()

# Add Truckmate RDP shortcut
$wshshell = New-Object -ComObject WScript.Shell
$lnk = $wshshell.CreateShortcut("C:\Users\Public\Desktop\Truckmate.lnk")
$lnk.TargetPath = "%windir%\system32\mstsc.exe"
$lnk.Arguments = "/v:ncs-rds01.ad.ncss.net"
$lnk.Description = "Truckmate"
$rdpFile = @"
screen mode id:i:2
use multimon:i:0
desktopwidth:i:1920
desktopheight:i:1080
session bpp:i:32
winposstr:s:0,3,0,0,800,600
compression:i:1
keyboardhook:i:2
audiocapturemode:i:0
videoplaybackmode:i:1
connection type:i:7
networkautodetect:i:1
bandwidthautodetect:i:1
displayconnectionbar:i:1
enableworkspacereconnect:i:0
disable wallpaper:i:0
allow font smoothing:i:0
allow desktop composition:i:0
disable full window drag:i:1
disable menu anims:i:1
disable themes:i:0
disable cursor setting:i:0
bitmapcachepersistenable:i:1
full address:s:Truckmate.ad.ncss.net
audiomode:i:0
redirectprinters:i:1
redirectcomports:i:0
redirectsmartcards:i:1
redirectclipboard:i:1
redirectposdevices:i:0
autoreconnection enabled:i:1
authentication level:i:2
prompt for credentials:i:0
negotiate security layer:i:1
remoteapplicationmode:i:0
alternate shell:s:
shell working directory:s:
gatewayhostname:s:
gatewayusagemethod:i:4
gatewaycredentialssource:i:4
gatewayprofileusagemethod:i:0
promptcredentialonce:i:0
gatewaybrokeringtype:i:0
use redirection server name:i:0
rdgiskdcproxy:i:0
kdcproxyname:s:
drivestoredirect:s:I:\;S:\;X:\;Y:\;
"@
$rdpFile | Out-File C:\ProgramData\Desktop\Truckmate.rdp

# Add Revenova Web shortcut
$new_object_2 = New-Object -ComObject WScript.Shell
$destination = $new_object_2.SpecialFolders.Item("AllUsersDesktop")
$source_path = Join-Path -Path $destination -ChildPath "\\Revenova.url"
$source = $new_object_2.CreateShortcut($source_path)
$source.TargetPath = "https://ucgroup.my.salesforce.com/"
$source.Save()

