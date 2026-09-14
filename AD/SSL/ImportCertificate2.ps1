# Define variables
$RemoteComputer = "AZUPRDBLD01"    # Replace with the remote computer's name or IP
$PfxFilePath = "C:\Temp\DigiCert\2025\2026WC.pfx" # Local path to the .pfx file
$PfxPassphrase = "W1ldC@rd!"          # Password for the .pfx file
$CertStoreName = "LocalMachine\My"         # Target certificate store (Personal for Computer Account)

# Copy the PFX file to the remote computer
$RemotePfxPath = "\\$RemoteComputer\C$\Temp\2026WC.pfx"
Copy-Item -Path $PfxFilePath -Destination $RemotePfxPath -Force

# Install the certificate on the remote computer
Invoke-Command -ComputerName $RemoteComputer -ScriptBlock {
    param ($PfxPath, $Passphrase, $StoreName)

    # Load the PFX file with the provided passphrase
    $SecurePass = ConvertTo-SecureString -String $Passphrase -AsPlainText -Force
    $Cert = New-Object System.Security.Cryptography.X509Certificates.X509Certificate2
    $Cert.Import($PfxPath, $SecurePass, [System.Security.Cryptography.X509Certificates.X509KeyStorageFlags]::MachineKeySet)

    # Open the target certificate store
    $Store = New-Object System.Security.Cryptography.X509Certificates.X509Store($StoreName, [System.Security.Cryptography.X509Certificates.StoreLocation]::LocalMachine)
    $Store.Open("ReadWrite")

    # Add the certificate to the store
    $Store.Add($Cert)
    $Store.Close()

    Write-Output "Certificate imported successfully to $StoreName"
} -ArgumentList $RemotePfxPath, $PfxPassphrase, $CertStoreName

# Clean up the remote file
# Remove-Item -Path $RemotePfxPath -Force