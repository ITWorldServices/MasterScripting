# Connect to Exchange Online
# Commercial
# Connect-ExchangeOnline
# or
# GCC High
Connect-ExchangeOnline -ExchangeEnvironmentName O365USGovGCCHigh

# Import CSV
$DLs = Import-Csv "C:\Temp\apm-dls.csv"

# Loop through each row in the CSV and create DL
foreach ($dl in $DLs) {
    $params = @{
        DisplayName        = $dl.DisplayName
        Name               = $dl.Alias
        PrimarySmtpAddress = $dl.PrimarySmtpAddress
    }
    New-DistributionGroup @params
}

# Disconnect when done
Disconnect-ExchangeOnline
