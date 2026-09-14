#region printer list
$availablePrinters = @(
    [pscustomobject]@{
        SharedID   = '062f4402-83c2-48ce-9acb-31f2951daee2'
        SharedName = 'UC-NA-HQ-Accounting Printer'
        IsDefault  = $null
    }
    [pscustomobject]@{
        SharedID   = '0efa396c-2948-4f25-9397-e763a597ceca'
        SharedName = 'UC-NA-HQ-Dispatch Pritner'
        IsDefault  = $null
    }
    [pscustomobject]@{
        SharedID   = '46a2ceaf-586f-4a7c-bbe8-fe33fb982ffa'
        SharedName = 'UC-NA-HQ-HR Printer'
        IsDefault  = $null
    }
    [pscustomobject]@{
        SharedID   = 'b3eac6a7-2282-4acf-89f0-d68f0e74b2d9'
        SharedName = 'UC-NA-HQ-Admin Printer'
        IsDefault  = $null
    }
    [pscustomobject]@{
        SharedID   = 'beb3607f-b7d0-461a-ab14-e0a64ddeca66'
        SharedName = 'UC-NA-HQ-Warehouse Printer'
        IsDefault  = $null
    }
)
#endregion
try {
    $configurationPath = "$env:appdata\UniversalPrintPrinterProvisioning\Configuration"
    if (!(Test-Path $configurationPath -ErrorAction SilentlyContinue)) {
        New-Item $configurationPath -ItemType Directory -Force | Out-Null
    }
    $printCfg = ($availablePrinters | ConvertTo-Csv -NoTypeInformation | ForEach-Object { $_ -replace '"', "" } ) -join [System.Environment]::NewLine
    $printCfg | Out-File "$configurationPath\printers.csv" -Encoding ascii -NoNewline
    Start-Process "${env:ProgramFiles(x86)}\UniversalPrintPrinterProvisioning\Exe\UPPrinterInstaller.exe" -Wait -WindowStyle Hidden
}
catch {
    $errorMsg = $_.Exception.Message
}