$agentInstallPath = Join-Path ${Env:ProgramFiles(x86)} "Beanywhere Support Express\GetSupportService_N-central"
if ($env:PROCESSOR_ARCHITECTURE -eq "x86") {
    $agentInstallPath = Join-Path ${Env:ProgramFiles} "Beanywhere Support Express\GetSupportService_N-central"
}

$AgentBinaryPath = Join-Path $agentInstallPath "BASupSrvc.exe"
$UpdaterBinaryPath = Join-Path $agentInstallPath "BASupSrvcUpdater.exe"
$AgentServiceName = "BASupportExpressStandaloneService_N_Central"
$UpdaterServiceName = "BASupportExpressSrvcUpdater_N_Central"
$InstallLockFilePath = Join-Path $agentInstallPath "__installing.lock"
$UnInstallLockFilePath = Join-Path $agentInstallPath "__uninstalling.lock"
$serviceNotRunningGuardInterval = 10
$lockFileAgeThresholdMinutes = 10

function CheckFileSignature  {
    param (
        [string]$FilePath
    )

    $result = $false

    try {

        $signature = Get-AuthenticodeSignature -FilePath $FilePath

        if ($signature.Status -eq "Valid") {

            if ($signature.SignerCertificate.Subject -eq "CN=N-ABLE TECHNOLOGIES LTD, O=N-ABLE TECHNOLOGIES LTD, L=Dundee, C=GB") {
                Write-Host "The file has a valid signature."
                $result = $true
            } else {
                Write-Error "The file has a valid signature but is not signed by N-able."
            }

        } else {
            Write-Error "The file does not have a valid signature."
        }

    } catch {
        Write-Error "Error: Unable to retrieve signature information for the file."
    }

    return $result

}

function FetchAndReinstall {

    $validRequest = $false
    try {

        $remoteJsonUrl = "https://swi-rc.cdn-sw.net/n-central/updates/json/TakeControlCheckAndReInstall.json"

        $jsonContent = Invoke-RestMethod -Uri $remoteJsonUrl
        $validRequest = $true

    }
    catch {
        Write-Error "Exception occurred while retrieving the remote json file."
    }

    if ($validRequest) { 
      
        try {

            $Url = $jsonContent.url;
            $ExpectedHash = $jsonContent.expected_hash
            $ExpectedSize = $jsonContent.expected_size

        }
        catch {
            Write-Error "Exception occurred while parsing the remote json file."
            $validRequest = $false
        }
 
        if (($Url -ne "") -and ($ExpectedHash -ne "") -and ($validRequest)) {

            $FilePath = Join-Path $env:TEMP "MSPA4NCentralInstaller.exe"
            $Parameters = "/S /R /L"

            Remove-Item -Path $FilePath -ErrorAction SilentlyContinue

            Write-Host "Fetching Take Control agent binary from '$Url' to '$FilePath'."
            Invoke-WebRequest -Uri $Url -OutFile $FilePath

            Write-Host "Verifying the hash of the downloaded file."
            $ActualHash = (Get-FileHash -Path $FilePath -Algorithm SHA256).Hash

            $ActualSize = (Get-Item -Path $FilePath).Length

            if ($ExpectedSize -ne $ActualSize) {
                Write-Error "The file size does not match the expected size. Exiting..."
            } 
            elseif ($ExpectedHash -ne $ActualHash) {
                Write-Error "The file hash does not match the expected hash. Exiting..."
            }
            elseif (-not (CheckFileSignature($FilePath))) {
                Write-Error "The file signature is not valid. Exiting..."   
            }
            else {
                Write-Host "The file size and hash match the expected values and the signature is correct. Running agent installer..."
                Start-Process -FilePath $FilePath -ArgumentList $Parameters -Wait
            }

            Remove-Item -Path $FilePath

        }
        else {
            Write-Error "Empty URL or expected_hash."
        }

    }
    else {
        Write-Error "Unable to retrieve the remote json file."
    }

    Exit    

}

function CheckLockFileAndReInstall {

    $lockExists = $false

    if (Test-Path -Path $InstallLockFilePath) {
        $installLockFileCreationTime = (Get-Item -Path $InstallLockFilePath).CreationTime
        $ageMinutes = (Get-Date) - $installLockFileCreationTime
        if ($ageMinutes.TotalMinutes -lt $lockFileAgeThresholdMinutes) {
            Write-Host "The lock file '$InstallLockFilePath' is newer than $lockFileAgeThresholdMinutes minutes. Exiting..."
            $lockExists = $true
        }
        else {
            Write-Host "The lock file '$InstallLockFilePath' is older than $lockFileAgeThresholdMinutes minutes."
        }
    }
    elseif (Test-Path -Path $UnInstallLockFilePath) {
        $unInstallLockFileCreationTime = (Get-Item -Path $UnInstallLockFilePath).CreationTime
        $ageMinutes = (Get-Date) - $unInstallLockFileCreationTime
        if ($ageMinutes.TotalMinutes -gt $lockFileAgeThresholdMinutes) {
            Write-Host "The lock file '$UnInstallLockFilePath' is newer than $lockFileAgeThresholdMinutes minutes. Exiting..."
            $lockExists = $true
        }
        else {
            Write-Host "The lock file '$UnInstallLockFilePath' is older than $lockFileAgeThresholdMinutes minutes."
        }
    }
    else {
        Write-Host "The lock file '$InstallLockFilePath' does not exist."
    }

    if ($lockExists -eq $false) {
        FetchAndReinstall
    }

    Exit

}

function WaitForServiceToStart {
    param (
        [Parameter(Mandatory = $true)]
        [string]$serviceName,

        [Parameter(Mandatory = $true)]
        [int]$waitTimeInMinutes
    )

    $endTime = (Get-Date).AddMinutes($waitTimeInMinutes)

    while ((Get-Date) -lt $endTime) {
        $service = Get-Service -Name $serviceName -ErrorAction SilentlyContinue

        if ($service -ne $null -and $service.Status -eq 'Running') {
            Write-Host "Service '$serviceName' has started."
            return $true
        }

        Start-Sleep -Seconds 5
    }

    Write-Host "Service '$serviceName' did not start within the specified wait time."
    return $false

}

if (-NOT ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")) {
    Write-Host "This script must be run with Administrator privileges."
    Exit
} 

# add a command line argument to force re-installation
if ($args.Count -eq 1 -and $args[0] -eq "-force") {

    Write-Host "Forcing re-installation..."
    CheckLockFileAndReInstall

}

if ((-not (Test-Path -Path $AgentBinaryPath)) -or (-not (Test-Path -Path $UpdaterBinaryPath))) {

    Write-Host "The Take Control agent binaries do not found..."
    CheckLockFileAndReInstall

} else {

    Write-Host "The Take Control agent binaries found..."

}

$agentService = Get-Service -Name $AgentServiceName -ErrorAction SilentlyContinue
if (-not $agentService) {

    Write-Host "The service '$AgentServiceName' is not registered..."
    CheckLockFileAndReInstall

} else {

    Write-Host "The service '$AgentServiceName' is registered..."

}

$updaterService = Get-Service -Name $UpdaterServiceName -ErrorAction SilentlyContinue
if (-not $updaterService) {

    Write-Host "The service '$UpdaterServiceName' is not registered."
    CheckLockFileAndReInstall

} else {

    Write-Host "The service '$UpdaterServiceName' is registered..."

}

if ($agentService.Status -ne "Running") {

    Write-Host "The service '$AgentServiceName' is not running... Waiting..."
    $agentServiceStarted = WaitForServiceToStart -serviceName $AgentServiceName -waitTimeInMinutes $serviceNotRunningGuardInterval
    if ($agentServiceStarted -eq $false) {
        Write-Host "The service '$AgentServiceName' is still not running... Re-Installing..."
        CheckLockFileAndReInstall
    }
    else {
        Write-Host "The service '$UpdaterServiceName' started... Skipping re-installation..."
    }

} else {

    Write-Host "The service '$AgentServiceName' is running..."

}

if ($updaterService.Status -ne "Running") {  

    Write-Host "The service '$UpdaterServiceName' is not running... Waiting..."
    $updaterServiceStarted = WaitForServiceToStart -serviceName $UpdaterServiceName -waitTimeInMinutes $serviceNotRunningGuardInterval
    if ($updaterServiceStarted -eq $false) {
        Write-Host "The service '$UpdaterServiceName' is still not running... Re-Installing..."
        CheckLockFileAndReInstall
    }
    else {
        Write-Host "The service '$UpdaterServiceName' started... Skipping re-installation..."
    }
 
} else {

    Write-Host "The service '$UpdaterServiceName' is running..."

}

Write-Host "All Take Control services are running... Exiting..."
# SIG # Begin signature block
# MIIjbAYJKoZIhvcNAQcCoIIjXTCCI1kCAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCDtywe+5djEyS+N
# eNXdS22Ki1dGZFnGUEqMBg2wW8DweaCCHWUwggUkMIIEDKADAgECAhAPH127GmWL
# aPvoHx8cY/CUMA0GCSqGSIb3DQEBCwUAMHIxCzAJBgNVBAYTAlVTMRUwEwYDVQQK
# EwxEaWdpQ2VydCBJbmMxGTAXBgNVBAsTEHd3dy5kaWdpY2VydC5jb20xMTAvBgNV
# BAMTKERpZ2lDZXJ0IFNIQTIgQXNzdXJlZCBJRCBDb2RlIFNpZ25pbmcgQ0EwHhcN
# MjEwNTI3MDAwMDAwWhcNMjQwNTMwMjM1OTU5WjBiMQswCQYDVQQGEwJHQjEPMA0G
# A1UEBxMGRHVuZGVlMSAwHgYDVQQKExdOLUFCTEUgVEVDSE5PTE9HSUVTIExURDEg
# MB4GA1UEAxMXTi1BQkxFIFRFQ0hOT0xPR0lFUyBMVEQwggEiMA0GCSqGSIb3DQEB
# AQUAA4IBDwAwggEKAoIBAQC/QTHwo4U3HBy/P+WdRniKXkpUBGM3z3jac892rkaU
# gJ2xjbecRSTDUgdx3P81cHkNAzwYwerLJcmLKl7MGZdaNBsB3W0i7i0w3hy5pjYs
# PQbe+hid03HNEUoDJEwZnBwlvRi6MPWGRPnU6IJmo5WdKQgKwRzOkAYjKq1Pdlrc
# qeNlsXjZXeMKujGLNrFJG1DcQ2lSuaDCDe0DvWiGYhXwz8PbeyDO2rsI9MqGk9DB
# FTywYCtmffJdyhyKb9twMWmhYCMFvU6LjDcnodAF8Vr1QXomXsiRcIze2l4eMOzN
# j/s9nM//tIVLvubxk4yqBzWj7Fnt5PhgcDcnRzF2Zfe5AgMBAAGjggHEMIIBwDAf
# BgNVHSMEGDAWgBRaxLl7KgqjpepxA8Bg+S32ZXUOWDAdBgNVHQ4EFgQU+rp2Tf6I
# w4DTidUzDAdflWwQrE4wDgYDVR0PAQH/BAQDAgeAMBMGA1UdJQQMMAoGCCsGAQUF
# BwMDMHcGA1UdHwRwMG4wNaAzoDGGL2h0dHA6Ly9jcmwzLmRpZ2ljZXJ0LmNvbS9z
# aGEyLWFzc3VyZWQtY3MtZzEuY3JsMDWgM6Axhi9odHRwOi8vY3JsNC5kaWdpY2Vy
# dC5jb20vc2hhMi1hc3N1cmVkLWNzLWcxLmNybDBLBgNVHSAERDBCMDYGCWCGSAGG
# /WwDATApMCcGCCsGAQUFBwIBFhtodHRwOi8vd3d3LmRpZ2ljZXJ0LmNvbS9DUFMw
# CAYGZ4EMAQQBMIGEBggrBgEFBQcBAQR4MHYwJAYIKwYBBQUHMAGGGGh0dHA6Ly9v
# Y3NwLmRpZ2ljZXJ0LmNvbTBOBggrBgEFBQcwAoZCaHR0cDovL2NhY2VydHMuZGln
# aWNlcnQuY29tL0RpZ2lDZXJ0U0hBMkFzc3VyZWRJRENvZGVTaWduaW5nQ0EuY3J0
# MAwGA1UdEwEB/wQCMAAwDQYJKoZIhvcNAQELBQADggEBAOJ6ufnLgmDb7nYghoqD
# XJjI5LFIzrE64QLOg1WWRqcaipYp39BgEGYfQi+oEGD26z+4jsdyyba6D6APkcrk
# LkYxuKu5e2TXnsQMeyweQHc0IymHDLdr8tUcJPsVoPDnmzgDeU0uih4F6SRgGdxf
# k+GsEls3efL51XuynpVCZwSOwLDZ0ZHUZ8Pn9hXjS+HsllgWysn8f84a6Xz2w2gS
# G8fd47bhSGyq5bTrKcXnmAiylCxoYFQqObIQgWE1w0/AAhUsal36zH14Wat95hZ5
# rnVQpmovf/tUNQKR3jyk68w3mNNAkItbm1TJ6JAuslY0VBZvUaiahqiAL6zn3NtY
# 8w8wggUwMIIEGKADAgECAhAECRgbX9W7ZnVTQ7VvlVAIMA0GCSqGSIb3DQEBCwUA
# MGUxCzAJBgNVBAYTAlVTMRUwEwYDVQQKEwxEaWdpQ2VydCBJbmMxGTAXBgNVBAsT
# EHd3dy5kaWdpY2VydC5jb20xJDAiBgNVBAMTG0RpZ2lDZXJ0IEFzc3VyZWQgSUQg
# Um9vdCBDQTAeFw0xMzEwMjIxMjAwMDBaFw0yODEwMjIxMjAwMDBaMHIxCzAJBgNV
# BAYTAlVTMRUwEwYDVQQKEwxEaWdpQ2VydCBJbmMxGTAXBgNVBAsTEHd3dy5kaWdp
# Y2VydC5jb20xMTAvBgNVBAMTKERpZ2lDZXJ0IFNIQTIgQXNzdXJlZCBJRCBDb2Rl
# IFNpZ25pbmcgQ0EwggEiMA0GCSqGSIb3DQEBAQUAA4IBDwAwggEKAoIBAQD407Mc
# fw4Rr2d3B9MLMUkZz9D7RZmxOttE9X/lqJ3bMtdx6nadBS63j/qSQ8Cl+YnUNxnX
# tqrwnIal2CWsDnkoOn7p0WfTxvspJ8fTeyOU5JEjlpB3gvmhhCNmElQzUHSxKCa7
# JGnCwlLyFGeKiUXULaGj6YgsIJWuHEqHCN8M9eJNYBi+qsSyrnAxZjNxPqxwoqvO
# f+l8y5Kh5TsxHM/q8grkV7tKtel05iv+bMt+dDk2DZDv5LVOpKnqagqrhPOsZ061
# xPeM0SAlI+sIZD5SlsHyDxL0xY4PwaLoLFH3c7y9hbFig3NBggfkOItqcyDQD2Rz
# PJ6fpjOp/RnfJZPRAgMBAAGjggHNMIIByTASBgNVHRMBAf8ECDAGAQH/AgEAMA4G
# A1UdDwEB/wQEAwIBhjATBgNVHSUEDDAKBggrBgEFBQcDAzB5BggrBgEFBQcBAQRt
# MGswJAYIKwYBBQUHMAGGGGh0dHA6Ly9vY3NwLmRpZ2ljZXJ0LmNvbTBDBggrBgEF
# BQcwAoY3aHR0cDovL2NhY2VydHMuZGlnaWNlcnQuY29tL0RpZ2lDZXJ0QXNzdXJl
# ZElEUm9vdENBLmNydDCBgQYDVR0fBHoweDA6oDigNoY0aHR0cDovL2NybDQuZGln
# aWNlcnQuY29tL0RpZ2lDZXJ0QXNzdXJlZElEUm9vdENBLmNybDA6oDigNoY0aHR0
# cDovL2NybDMuZGlnaWNlcnQuY29tL0RpZ2lDZXJ0QXNzdXJlZElEUm9vdENBLmNy
# bDBPBgNVHSAESDBGMDgGCmCGSAGG/WwAAgQwKjAoBggrBgEFBQcCARYcaHR0cHM6
# Ly93d3cuZGlnaWNlcnQuY29tL0NQUzAKBghghkgBhv1sAzAdBgNVHQ4EFgQUWsS5
# eyoKo6XqcQPAYPkt9mV1DlgwHwYDVR0jBBgwFoAUReuir/SSy4IxLVGLp6chnfNt
# yA8wDQYJKoZIhvcNAQELBQADggEBAD7sDVoks/Mi0RXILHwlKXaoHV0cLToaxO8w
# Ydd+C2D9wz0PxK+L/e8q3yBVN7Dh9tGSdQ9RtG6ljlriXiSBThCk7j9xjmMOE0ut
# 119EefM2FAaK95xGTlz/kLEbBw6RFfu6r7VRwo0kriTGxycqoSkoGjpxKAI8LpGj
# wCUR4pwUR6F6aGivm6dcIFzZcbEMj7uo+MUSaJ/PQMtARKUT8OZkDCUIQjKyNook
# Av4vcn4c10lFluhZHen6dGRrsutmQ9qzsIzV6Q3d9gEgzpkxYz0IGhizgZtPxpMQ
# BvwHgfqL2vmCSfdibqFT+hKUGIUukpHqaGxEMrJmoecYpJpkUe8wggWNMIIEdaAD
# AgECAhAOmxiO+dAt5+/bUOIIQBhaMA0GCSqGSIb3DQEBDAUAMGUxCzAJBgNVBAYT
# AlVTMRUwEwYDVQQKEwxEaWdpQ2VydCBJbmMxGTAXBgNVBAsTEHd3dy5kaWdpY2Vy
# dC5jb20xJDAiBgNVBAMTG0RpZ2lDZXJ0IEFzc3VyZWQgSUQgUm9vdCBDQTAeFw0y
# MjA4MDEwMDAwMDBaFw0zMTExMDkyMzU5NTlaMGIxCzAJBgNVBAYTAlVTMRUwEwYD
# VQQKEwxEaWdpQ2VydCBJbmMxGTAXBgNVBAsTEHd3dy5kaWdpY2VydC5jb20xITAf
# BgNVBAMTGERpZ2lDZXJ0IFRydXN0ZWQgUm9vdCBHNDCCAiIwDQYJKoZIhvcNAQEB
# BQADggIPADCCAgoCggIBAL/mkHNo3rvkXUo8MCIwaTPswqclLskhPfKK2FnC4Smn
# PVirdprNrnsbhA3EMB/zG6Q4FutWxpdtHauyefLKEdLkX9YFPFIPUh/GnhWlfr6f
# qVcWWVVyr2iTcMKyunWZanMylNEQRBAu34LzB4TmdDttceItDBvuINXJIB1jKS3O
# 7F5OyJP4IWGbNOsFxl7sWxq868nPzaw0QF+xembud8hIqGZXV59UWI4MK7dPpzDZ
# Vu7Ke13jrclPXuU15zHL2pNe3I6PgNq2kZhAkHnDeMe2scS1ahg4AxCN2NQ3pC4F
# fYj1gj4QkXCrVYJBMtfbBHMqbpEBfCFM1LyuGwN1XXhm2ToxRJozQL8I11pJpMLm
# qaBn3aQnvKFPObURWBf3JFxGj2T3wWmIdph2PVldQnaHiZdpekjw4KISG2aadMre
# Sx7nDmOu5tTvkpI6nj3cAORFJYm2mkQZK37AlLTSYW3rM9nF30sEAMx9HJXDj/ch
# srIRt7t/8tWMcCxBYKqxYxhElRp2Yn72gLD76GSmM9GJB+G9t+ZDpBi4pncB4Q+U
# DCEdslQpJYls5Q5SUUd0viastkF13nqsX40/ybzTQRESW+UQUOsxxcpyFiIJ33xM
# dT9j7CFfxCBRa2+xq4aLT8LWRV+dIPyhHsXAj6KxfgommfXkaS+YHS312amyHeUb
# AgMBAAGjggE6MIIBNjAPBgNVHRMBAf8EBTADAQH/MB0GA1UdDgQWBBTs1+OC0nFd
# ZEzfLmc/57qYrhwPTzAfBgNVHSMEGDAWgBRF66Kv9JLLgjEtUYunpyGd823IDzAO
# BgNVHQ8BAf8EBAMCAYYweQYIKwYBBQUHAQEEbTBrMCQGCCsGAQUFBzABhhhodHRw
# Oi8vb2NzcC5kaWdpY2VydC5jb20wQwYIKwYBBQUHMAKGN2h0dHA6Ly9jYWNlcnRz
# LmRpZ2ljZXJ0LmNvbS9EaWdpQ2VydEFzc3VyZWRJRFJvb3RDQS5jcnQwRQYDVR0f
# BD4wPDA6oDigNoY0aHR0cDovL2NybDMuZGlnaWNlcnQuY29tL0RpZ2lDZXJ0QXNz
# dXJlZElEUm9vdENBLmNybDARBgNVHSAECjAIMAYGBFUdIAAwDQYJKoZIhvcNAQEM
# BQADggEBAHCgv0NcVec4X6CjdBs9thbX979XB72arKGHLOyFXqkauyL4hxppVCLt
# pIh3bb0aFPQTSnovLbc47/T/gLn4offyct4kvFIDyE7QKt76LVbP+fT3rDB6mouy
# XtTP0UNEm0Mh65ZyoUi0mcudT6cGAxN3J0TU53/oWajwvy8LpunyNDzs9wPHh6jS
# TEAZNUZqaVSwuKFWjuyk1T3osdz9HNj0d1pcVIxv76FQPfx2CWiEn2/K2yCNNWAc
# AgPLILCsWKAOQGPFmCLBsln1VWvPJ6tsds5vIy30fnFqI2si/xK4VC0nftg62fC2
# h5b9W9FcrBjDTZ9ztwGpn1eqXijiuZQwggauMIIElqADAgECAhAHNje3JFR82Ees
# /ShmKl5bMA0GCSqGSIb3DQEBCwUAMGIxCzAJBgNVBAYTAlVTMRUwEwYDVQQKEwxE
# aWdpQ2VydCBJbmMxGTAXBgNVBAsTEHd3dy5kaWdpY2VydC5jb20xITAfBgNVBAMT
# GERpZ2lDZXJ0IFRydXN0ZWQgUm9vdCBHNDAeFw0yMjAzMjMwMDAwMDBaFw0zNzAz
# MjIyMzU5NTlaMGMxCzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5j
# LjE7MDkGA1UEAxMyRGlnaUNlcnQgVHJ1c3RlZCBHNCBSU0E0MDk2IFNIQTI1NiBU
# aW1lU3RhbXBpbmcgQ0EwggIiMA0GCSqGSIb3DQEBAQUAA4ICDwAwggIKAoICAQDG
# hjUGSbPBPXJJUVXHJQPE8pE3qZdRodbSg9GeTKJtoLDMg/la9hGhRBVCX6SI82j6
# ffOciQt/nR+eDzMfUBMLJnOWbfhXqAJ9/UO0hNoR8XOxs+4rgISKIhjf69o9xBd/
# qxkrPkLcZ47qUT3w1lbU5ygt69OxtXXnHwZljZQp09nsad/ZkIdGAHvbREGJ3Hxq
# V3rwN3mfXazL6IRktFLydkf3YYMZ3V+0VAshaG43IbtArF+y3kp9zvU5EmfvDqVj
# bOSmxR3NNg1c1eYbqMFkdECnwHLFuk4fsbVYTXn+149zk6wsOeKlSNbwsDETqVcp
# licu9Yemj052FVUmcJgmf6AaRyBD40NjgHt1biclkJg6OBGz9vae5jtb7IHeIhTZ
# girHkr+g3uM+onP65x9abJTyUpURK1h0QCirc0PO30qhHGs4xSnzyqqWc0Jon7ZG
# s506o9UD4L/wojzKQtwYSH8UNM/STKvvmz3+DrhkKvp1KCRB7UK/BZxmSVJQ9FHz
# NklNiyDSLFc1eSuo80VgvCONWPfcYd6T/jnA+bIwpUzX6ZhKWD7TA4j+s4/TXkt2
# ElGTyYwMO1uKIqjBJgj5FBASA31fI7tk42PgpuE+9sJ0sj8eCXbsq11GdeJgo1gJ
# ASgADoRU7s7pXcheMBK9Rp6103a50g5rmQzSM7TNsQIDAQABo4IBXTCCAVkwEgYD
# VR0TAQH/BAgwBgEB/wIBADAdBgNVHQ4EFgQUuhbZbU2FL3MpdpovdYxqII+eyG8w
# HwYDVR0jBBgwFoAU7NfjgtJxXWRM3y5nP+e6mK4cD08wDgYDVR0PAQH/BAQDAgGG
# MBMGA1UdJQQMMAoGCCsGAQUFBwMIMHcGCCsGAQUFBwEBBGswaTAkBggrBgEFBQcw
# AYYYaHR0cDovL29jc3AuZGlnaWNlcnQuY29tMEEGCCsGAQUFBzAChjVodHRwOi8v
# Y2FjZXJ0cy5kaWdpY2VydC5jb20vRGlnaUNlcnRUcnVzdGVkUm9vdEc0LmNydDBD
# BgNVHR8EPDA6MDigNqA0hjJodHRwOi8vY3JsMy5kaWdpY2VydC5jb20vRGlnaUNl
# cnRUcnVzdGVkUm9vdEc0LmNybDAgBgNVHSAEGTAXMAgGBmeBDAEEAjALBglghkgB
# hv1sBwEwDQYJKoZIhvcNAQELBQADggIBAH1ZjsCTtm+YqUQiAX5m1tghQuGwGC4Q
# TRPPMFPOvxj7x1Bd4ksp+3CKDaopafxpwc8dB+k+YMjYC+VcW9dth/qEICU0MWfN
# thKWb8RQTGIdDAiCqBa9qVbPFXONASIlzpVpP0d3+3J0FNf/q0+KLHqrhc1DX+1g
# tqpPkWaeLJ7giqzl/Yy8ZCaHbJK9nXzQcAp876i8dU+6WvepELJd6f8oVInw1Ypx
# dmXazPByoyP6wCeCRK6ZJxurJB4mwbfeKuv2nrF5mYGjVoarCkXJ38SNoOeY+/um
# nXKvxMfBwWpx2cYTgAnEtp/Nh4cku0+jSbl3ZpHxcpzpSwJSpzd+k1OsOx0ISQ+U
# zTl63f8lY5knLD0/a6fxZsNBzU+2QJshIUDQtxMkzdwdeDrknq3lNHGS1yZr5Dhz
# q6YBT70/O3itTK37xJV77QpfMzmHQXh6OOmc4d0j/R0o08f56PGYX/sr2H7yRp11
# LB4nLCbbbxV7HhmLNriT1ObyF5lZynDwN7+YAN8gFk8n+2BnFqFmut1VwDophrCY
# oCvtlUG3OtUVmDG0YgkPCr2B2RP+v6TR81fZvAT6gt4y3wSJ8ADNXcL50CN/AAvk
# dgIm2fBldkKmKYcJRyvmfxqkhQ/8mJb2VVQrH4D6wPIOK+XW+6kvRBVK5xMOHds3
# OBqhK/bt1nz8MIIGwjCCBKqgAwIBAgIQBUSv85SdCDmmv9s/X+VhFjANBgkqhkiG
# 9w0BAQsFADBjMQswCQYDVQQGEwJVUzEXMBUGA1UEChMORGlnaUNlcnQsIEluYy4x
# OzA5BgNVBAMTMkRpZ2lDZXJ0IFRydXN0ZWQgRzQgUlNBNDA5NiBTSEEyNTYgVGlt
# ZVN0YW1waW5nIENBMB4XDTIzMDcxNDAwMDAwMFoXDTM0MTAxMzIzNTk1OVowSDEL
# MAkGA1UEBhMCVVMxFzAVBgNVBAoTDkRpZ2lDZXJ0LCBJbmMuMSAwHgYDVQQDExdE
# aWdpQ2VydCBUaW1lc3RhbXAgMjAyMzCCAiIwDQYJKoZIhvcNAQEBBQADggIPADCC
# AgoCggIBAKNTRYcdg45brD5UsyPgz5/X5dLnXaEOCdwvSKOXejsqnGfcYhVYwamT
# EafNqrJq3RApih5iY2nTWJw1cb86l+uUUI8cIOrHmjsvlmbjaedp/lvD1isgHMGX
# lLSlUIHyz8sHpjBoyoNC2vx/CSSUpIIa2mq62DvKXd4ZGIX7ReoNYWyd/nFexAaa
# PPDFLnkPG2ZS48jWPl/aQ9OE9dDH9kgtXkV1lnX+3RChG4PBuOZSlbVH13gpOWvg
# eFmX40QrStWVzu8IF+qCZE3/I+PKhu60pCFkcOvV5aDaY7Mu6QXuqvYk9R28mxyy
# t1/f8O52fTGZZUdVnUokL6wrl76f5P17cz4y7lI0+9S769SgLDSb495uZBkHNwGR
# Dxy1Uc2qTGaDiGhiu7xBG3gZbeTZD+BYQfvYsSzhUa+0rRUGFOpiCBPTaR58ZE2d
# D9/O0V6MqqtQFcmzyrzXxDtoRKOlO0L9c33u3Qr/eTQQfqZcClhMAD6FaXXHg2TW
# dc2PEnZWpST618RrIbroHzSYLzrqawGw9/sqhux7UjipmAmhcbJsca8+uG+W1eEQ
# E/5hRwqM/vC2x9XH3mwk8L9CgsqgcT2ckpMEtGlwJw1Pt7U20clfCKRwo+wK8REu
# ZODLIivK8SgTIUlRfgZm0zu++uuRONhRB8qUt+JQofM604qDy0B7AgMBAAGjggGL
# MIIBhzAOBgNVHQ8BAf8EBAMCB4AwDAYDVR0TAQH/BAIwADAWBgNVHSUBAf8EDDAK
# BggrBgEFBQcDCDAgBgNVHSAEGTAXMAgGBmeBDAEEAjALBglghkgBhv1sBwEwHwYD
# VR0jBBgwFoAUuhbZbU2FL3MpdpovdYxqII+eyG8wHQYDVR0OBBYEFKW27xPn783Q
# ZKHVVqllMaPe1eNJMFoGA1UdHwRTMFEwT6BNoEuGSWh0dHA6Ly9jcmwzLmRpZ2lj
# ZXJ0LmNvbS9EaWdpQ2VydFRydXN0ZWRHNFJTQTQwOTZTSEEyNTZUaW1lU3RhbXBp
# bmdDQS5jcmwwgZAGCCsGAQUFBwEBBIGDMIGAMCQGCCsGAQUFBzABhhhodHRwOi8v
# b2NzcC5kaWdpY2VydC5jb20wWAYIKwYBBQUHMAKGTGh0dHA6Ly9jYWNlcnRzLmRp
# Z2ljZXJ0LmNvbS9EaWdpQ2VydFRydXN0ZWRHNFJTQTQwOTZTSEEyNTZUaW1lU3Rh
# bXBpbmdDQS5jcnQwDQYJKoZIhvcNAQELBQADggIBAIEa1t6gqbWYF7xwjU+KPGic
# 2CX/yyzkzepdIpLsjCICqbjPgKjZ5+PF7SaCinEvGN1Ott5s1+FgnCvt7T1Ijrhr
# unxdvcJhN2hJd6PrkKoS1yeF844ektrCQDifXcigLiV4JZ0qBXqEKZi2V3mP2yZW
# K7Dzp703DNiYdk9WuVLCtp04qYHnbUFcjGnRuSvExnvPnPp44pMadqJpddNQ5EQS
# viANnqlE0PjlSXcIWiHFtM+YlRpUurm8wWkZus8W8oM3NG6wQSbd3lqXTzON1I13
# fXVFoaVYJmoDRd7ZULVQjK9WvUzF4UbFKNOt50MAcN7MmJ4ZiQPq1JE3701S88lg
# IcRWR+3aEUuMMsOI5ljitts++V+wQtaP4xeR0arAVeOGv6wnLEHQmjNKqDbUuXKW
# fpd5OEhfysLcPTLfddY2Z1qJ+Panx+VPNTwAvb6cKmx5AdzaROY63jg7B145WPR8
# czFVoIARyxQMfq68/qTreWWqaNYiyjvrmoI1VygWy2nyMpqy0tg6uLFGhmu6F/3E
# d2wVbK6rr3M66ElGt9V/zLY4wNjsHPW2obhDLN9OTH0eaHDAdwrUAuBcYLso/zjl
# UlrWrBciI0707NMX+1Br/wd3H3GXREHJuEbTbDJ8WC9nR2XlG3O2mflrLAZG70Ee
# 8PBf4NvZrZCARK+AEEGKMYIFXTCCBVkCAQEwgYYwcjELMAkGA1UEBhMCVVMxFTAT
# BgNVBAoTDERpZ2lDZXJ0IEluYzEZMBcGA1UECxMQd3d3LmRpZ2ljZXJ0LmNvbTEx
# MC8GA1UEAxMoRGlnaUNlcnQgU0hBMiBBc3N1cmVkIElEIENvZGUgU2lnbmluZyBD
# QQIQDx9duxpli2j76B8fHGPwlDANBglghkgBZQMEAgEFAKCBhDAYBgorBgEEAYI3
# AgEMMQowCKACgAChAoAAMBkGCSqGSIb3DQEJAzEMBgorBgEEAYI3AgEEMBwGCisG
# AQQBgjcCAQsxDjAMBgorBgEEAYI3AgEVMC8GCSqGSIb3DQEJBDEiBCCQQPybASu8
# YZP+dnqmohbxM3MTOs/bsGHOhKjs4bdhiDANBgkqhkiG9w0BAQEFAASCAQAzb9pg
# dR2Gd4q81mmPY47ZENa+VJa1fsj1y9bYw2NjMPJ8FkwfObeDPCLrpvfdRNGWiupF
# 5jnOv6iF42xNa+CGnnC6Af6LJ+iBIdulhcYlsfGuQfomyPZIcE3kNZmhQfletU1N
# jnHsnOq/eh9riwAUISp5UT0kCJ9hURGnSckH99A75U+BuwTB4Mdz7M4y3/3cvkVw
# bjEdBoneVrynIPKAc08XrZwEOPbwdUs084Mv4ZA0c3iapuxBJGvbLpihPxtOGydh
# XHhEnPic3ceRgFspZ9++Eu56RFlCbQw8GenffxYVXciv83aHuj69SvHAmfXPmeYr
# qLA639rG2X/uATXeoYIDIDCCAxwGCSqGSIb3DQEJBjGCAw0wggMJAgEBMHcwYzEL
# MAkGA1UEBhMCVVMxFzAVBgNVBAoTDkRpZ2lDZXJ0LCBJbmMuMTswOQYDVQQDEzJE
# aWdpQ2VydCBUcnVzdGVkIEc0IFJTQTQwOTYgU0hBMjU2IFRpbWVTdGFtcGluZyBD
# QQIQBUSv85SdCDmmv9s/X+VhFjANBglghkgBZQMEAgEFAKBpMBgGCSqGSIb3DQEJ
# AzELBgkqhkiG9w0BBwEwHAYJKoZIhvcNAQkFMQ8XDTI0MDMyODE1MTAxOVowLwYJ
# KoZIhvcNAQkEMSIEIONhJEnp3ES6X5CPTLXRnZ8Q+040kmf/j3OMKRVUWQ5YMA0G
# CSqGSIb3DQEBAQUABIICAJNoX6/D6fjV4Yf/MteRHMSOfL3QzDCQg1FtXij/QguC
# ougQ3rsuqs1M5fUxdybZtMB6makAsMzt01q+tOTQuDKIMrPW264UkMvZHTzaXPCx
# VrnDVTwnRV/QW3QH4w6R/1A0n+EMjmrXjYEFHnT2DUWBcoZQKbojGiEIaq3STW95
# 0u1syu8NaYAh2nhV2kkalAN/lWjzWy9gYMLNpdVQFaX+3Hst/id0lVSHfmzTXa1u
# vZYqyVl//X7EAD/G6tdV20ONMZ0G9u+Os4kDE/OG409VXqMKCy0Jq3M0NsPQrLyb
# +040dfU6pYm/Ivdf7nnl/LCuj92dyt2a/QuYyAklvnsry/lrrbVyH38AlZTmlMRd
# j8Ce2ghq2qER0GmqqSv3Lz9bEdtZybPdHxfU5F8nra39sB7cqiBKVEzOV9uSovr6
# ehJHI0BtmCY5oSezYPXzQWL8wHh9097REZsA2t+glMhL36BfnzCJdUM79iSVGuIH
# d2S+fS/1HtEh6plCARRzrrQRH/LzCHTJkqSg1TcuUx63FRRIXemjMDj4ZoMjqPIk
# Q7gyw4nLCL//wSPbSYL0j0j4yQo9o1M7CEVxOHGFEKNzpP6S4qv5uJMC7NXRHZBF
# 8pflGXPY4KebPTg5XbxOnY3oaIEhmynKgvX2Q2IqpHHuBMt739j6NG+yLIn+1mVi
# SIG # End signature block
