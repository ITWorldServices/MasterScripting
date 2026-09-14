# First Set the PDC as the Time Keeper
# Execute this PS on the PDC
.\net.exe stop .\w32time.dll
.\w32tm.exe /config /syncfromflags:manualpeerlist:"0.us.pool.ntp.org,1.us.pool.ntp.org,2.us.pool.ntp.org,3.us.pool.ntp.org"
.\net.exe start .\w32time.dll
.\w32tm.exe /query /configuration
.\w32tm.exe /query /status
.\w32tm.exe /resync /rediscover