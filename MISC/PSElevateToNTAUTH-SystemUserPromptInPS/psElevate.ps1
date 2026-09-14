# Ensure the temp directory exists
if (-not (Test-Path "C:\temp")) {
    New-Item -Path "C:\temp" -ItemType Directory
}

# Download PsExec from Sysinternals
Invoke-WebRequest -Uri "https://live.sysinternals.com/PsExec.exe" -OutFile "C:\temp\PsExec.exe"

# Execute PowerShell with PsExec
Start-Process -FilePath "C:\temp\PsExec.exe" -ArgumentList "-i", "-s", "powershell.exe"
