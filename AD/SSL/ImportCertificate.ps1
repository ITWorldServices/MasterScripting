# Define variables
$RemoteComputer = "AZUPRDBLD01"    # Replace with the remote computer's name or IP
$CertFilePath = "\\d2lazfs01.file.core.windows.net\software\impact\DigiCert\2026WC.pfx" # Local path to the certificate file
$CertStoreName = "LocalMachine\My"        # Target certificate store (e.g., Personal store)

# Copy the certificate file to the remote computer
$RemoteCertPath = "\\$RemoteComputer\C$\Temp\"
Copy-Item -Path $CertFilePath -Destination $RemoteCertPath -Force

# Install the certificate on the remote computer
Invoke-Command -ComputerName $RemoteComputer -ScriptBlock {
    param ($CertPath, $StoreName)

    # Ensure the target folder exists
    $Store = New-Object System.Security.Cryptography.X509Certificates.X509Store($StoreName, [System.Security.Cryptography.X509Certificates.StoreLocation]::LocalMachine)
    $Store.Open('ReadWrite')

    # Import the certificate
    $Cert = New-Object System.Security.Cryptography.X509Certificates.X509Certificate2
    $Cert.Import($CertPath)
    $Store.Add($Cert)
    $Store.Close()

    Write-Output "Certificate imported successfully to $StoreName"
} -ArgumentList $RemoteCertPath, $CertStoreName

# Clean up the remote file
Remove-Item -Path $RemoteCertPath -Force