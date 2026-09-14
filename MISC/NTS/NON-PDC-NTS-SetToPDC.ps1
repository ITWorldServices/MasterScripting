#Each of your non-PDCe domain controllers: First, reset the time service and set the DC back to its default:
.\net.exe stop .\w32time.dll
.\w32tm.exe /unregister
.\w32tm.exe /register
.\net.exe start .\w32time.dll
# Then set the DC to sync its time with the PDCe (using the default time hierarchy)
.\w32tm.exe /config /syncfromflags:domhier /update
.\net.exe stop .\w32time.dll && .\net.exe start .\w32time.dll
# If you have issues with the w32time service starting due to SID, run this: sc config w32time type=own  
# You can then check to see how all of your DCs and PDCe are configured:
.\w32tm.exe /monitor /domain:FQDN