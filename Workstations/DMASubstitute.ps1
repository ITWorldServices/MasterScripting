#The following commands will reset Outlook to prompt user to create a new mail profile designed to replace the inadequate BitTitan DMA Tool 
#Create new backup directory
New-Item -Path C:\Temp\MailProfileBackup -ItemType Directory -Force

#Backup registry for all existing mail profiles to backup directory
reg export HKCU\SOFTWARE\Microsoft\Office\16.0\Outlook\Profiles c:\Temp\MailProfileBackup\allprofiles.reg

#Remove all mail profiles from registry
Remove-Item -Path HKCU:\SOFTWARE\Microsoft\Office\16.0\Outlook\Profiles\* -Recurse

#Upon completion you gain a full backup of all mail profiles reset Outlook to start new and keep existing ost data at its original location which allows you to restore if needed to old configuration by running the following command without the commented character

#reg import C:\Temp\MailProfileBackup\allprofiles.reg
