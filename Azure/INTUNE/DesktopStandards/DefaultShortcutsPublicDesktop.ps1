# Add APP1.UC.GROUP shortcut
$new_object_1 = New-Object -ComObject WScript.Shell
$destination = $new_object_1.SpecialFolders.Item("AllUsersDesktop")
$source_path = Join-Path -Path $destination -ChildPath "\\app1.url"
$source = $new_object_1.CreateShortcut($source_path)
$source.TargetPath = "https://app1.uc.group/RDWeb/Pages/en-US/login.aspx?ReturnUrl=/RDWeb/Pages/en-US/Default.aspx"
$source.Save()

# Add APP2.UC.GROUP shortcut
$new_object_2 = New-Object -ComObject WScript.Shell
$destination = $new_object_2.SpecialFolders.Item("AllUsersDesktop")
$source_path = Join-Path -Path $destination -ChildPath "\\app2.url"
$source = $new_object_2.CreateShortcut($source_path)
$source.TargetPath = "https://app2.uc.group/RDWeb/Pages/en-US/login.aspx?ReturnUrl=/RDWeb/Pages/en-US/Default.aspx"
$source.Save()

# Add Truckmate RDP shortcut
$wshshell = New-Object -ComObject WScript.Shell
$lnk = $wshshell.CreateShortcut("C:\Users\Public\Desktop\Truckmate.lnk")
$lnk.TargetPath = "%windir%\system32\mstsc.exe"
$lnk.Arguments = "/v:Truckmate.uc.group"
$lnk.Description = "Truckmate"
$lnk.Save()