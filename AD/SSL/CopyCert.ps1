# Define variables
$RemoteComputer = "AZUPRDBLD01"    # Replace with the remote computer's name or IP
$PfxFilePath = "C:\Temp\DigiCert\2025\2026WC.pfx" # Local path to the .pfx file
$PfxPassphrase = "W1ldC@rd!"          # Password for the .pfx file
$CertStoreName = "LocalMachine\My"         # Target certificate store (Personal for Computer Account)

# Copy the PFX file to the remote computer
$RemotePfxPath = "\\$RemoteComputer\C$\Temp\2026WC.pfx"
Copy-Item -Path $PfxFilePath -Destination $RemotePfxPath -Force