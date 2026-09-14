#Super simple script by Kelvin Aruin. 09/2022
#You can run this script from a computer list in a TXT file, or on each machine via INC or other RMM by using only the following:
#Get-ChildItem-Path 'C:\'-Include *.pst -File -Exclude 'C:\Windows\'-Recurse -ErrorAction SilentlyContinue |Select-ObjectDirectory,Name,Length,@{N="HostName";E={$hostname1.name}} | Export-Csv-append "\\PDC\PST\JacksonPSTInquiry.csv"-force -notypeinfo

#default configuration will run off a computers.txt computer list located in C:\scripts\ folder
#You will have to modify the \\HOSTNAME\SHARENAME\CSVNAME to your preffered directory such as \\SERVER3\PSTEXPORT\PSTExport.csv 
#Super simple otherwise!

$Computers = Get-Content -Path "C:\scripts\Computers.txt"
Foreach ($Computer in $Computers.Name)
{
Get-ChildItem-Path 'C:\'-Include *.pst -File -Exclude 'C:\Windows\'-Recurse -ErrorAction SilentlyContinue |Select-ObjectDirectory,Name,Length,@{N="HostName";E={$hostname1.name}} | 
Export-Csv-append "\\HOSTNAME\SHARENAME\CSVNAME.csv"-force -notypeinfo