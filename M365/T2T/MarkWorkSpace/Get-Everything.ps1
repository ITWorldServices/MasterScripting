# Set output paths
$timestamp = Get-Date -Format 'yyyyMMdd-HHmmss'
$OutputPath = "C:\temp"
$OutputFile = "$OutputPath\FullUserMailboxReport_$timestamp.csv"
$ErrorLog = "$OutputPath\MailboxExportErrors_$timestamp.log"
$OneDriveReportPath = "$OutputPath\OneDriveTempReport_$timestamp.csv"

# Connect to Microsoft Graph
Connect-MgGraph -Scopes "User.Read.All", "Organization.Read.All", "Reports.Read.All", "ReportSettings.ReadWrite.All"

# Connect to Exchange Online
Connect-ExchangeOnline

# Get all available SKUs (license types) for lookup
$allSkus = Get-MgSubscribedSku

# Try to get OneDrive usage report (30 days)
try {
    Get-MgReportOneDriveUsageAccountDetail -Period D30 -OutFile $OneDriveReportPath -ErrorAction Stop
    $oneDriveData = Import-Csv $OneDriveReportPath | Group-Object -AsHashTable -Property 'Owner Principal name'
} catch {
    Write-Error "Failed to get OneDrive report: $_"
    exit 1
}

# Get all Azure AD users with required properties, including AssignedLicenses
$users = Get-MgUser -All -Property `
    AssignedLicenses,GivenName,Surname,DisplayName,UserPrincipalName,JobTitle,Mail,Department,CompanyName,OfficeLocation,EmployeeId,MobilePhone,BusinessPhones,StreetAddress,City,PostalCode,State,Country,UserType,OnPremisesSyncEnabled,AccountEnabled,CreatedDateTime

# Prepare array for combined data
$combinedData = @()
$counter = 0

foreach ($user in $users) {
    $counter++
    try {
        #--- OneDrive Section ---
        $oneDriveProvisioned = $false
        $oneDriveUrl = "Not Provisioned"
        $lastActivity = "N/A"
        $fileCount = "0"
        $usedGB = "0.00"

        try {
            $drive = Get-MgUserDefaultDrive -UserId $user.UserPrincipalName -ErrorAction Stop
            $oneDriveProvisioned = $true
            $oneDriveUrl = $drive.WebUrl
            if ($oneDriveData.ContainsKey($user.UserPrincipalName)) {
                $reportEntry = $oneDriveData[$user.UserPrincipalName]
                $lastActivity = $reportEntry.'Last Activity Date'
                $fileCount = $reportEntry.'File Count'
                $usedGB = [math]::Round([decimal]$reportEntry.'Storage Used (Byte)' / 1GB, 2)
            }
        } catch {}

        #--- License Section ---
        $licenseAssigned = $false
        $licenseNames = ""
        if ($null -ne $user.AssignedLicenses -and $user.AssignedLicenses.Count -gt 0) {
            $licenseAssigned = $true
            $skuNames = @()
            foreach ($lic in $user.AssignedLicenses) {
                $sku = $allSkus | Where-Object { $_.SkuId -eq $lic.SkuId }
                if ($sku) {
                    $skuNames += $sku.SkuPartNumber
                }
            }
            $licenseNames = $skuNames -join ", "
        }

        #--- Mailbox Section ---
        $mailbox = $null
        $mailboxStats = $null
        $archiveStats = $null
        $mailboxHidden = "Unknown"
        $emailAddresses = ""
        $forwardingTarget = ""
        $forwardingBehavior = ""
        $mailboxForwarded = $false

        try {
            $mailbox = Get-Mailbox -Identity $user.UserPrincipalName -ErrorAction Stop
            $mailboxStats = Get-MailboxStatistics -Identity $user.UserPrincipalName -ErrorAction Stop
            if ($mailbox.ArchiveStatus -ne $null -and $mailbox.ArchiveStatus -ne 'None') {
                $archiveStats = Get-MailboxStatistics -Identity $mailbox.UserPrincipalName -Archive -ErrorAction SilentlyContinue
            }
            $mailboxHidden = if ($mailbox.HiddenFromAddressListsEnabled) { "Hidden" } else { "Visible" }

            # All email addresses (primary and aliases, SMTP only)
            $emailAddresses = ($mailbox.EmailAddresses | Where-Object { $_ -like "smtp:*" }) -replace "smtp:","" -join ", "

            # Forwarding info
            if ($mailbox.ForwardingSmtpAddress) {
                $forwardingTarget = $mailbox.ForwardingSmtpAddress -replace "smtp:",""
                $mailboxForwarded = $true
            } elseif ($mailbox.ForwardingAddress) {
                $recipient = Get-Recipient $mailbox.ForwardingAddress -ErrorAction SilentlyContinue
                if ($recipient) {
                    $forwardingTarget = $recipient.PrimarySmtpAddress
                    $mailboxForwarded = $true
                }
            }
            $forwardingBehavior = if ($mailboxForwarded) {
                if ($mailbox.DeliverToMailboxAndForward) { "Deliver to mailbox and forward" } else { "Forward only" }
            } else { "" }
        } catch {
            $mailbox = $null
            $mailboxStats = $null
            $archiveStats = $null
            $mailboxHidden = "No Mailbox"
            $emailAddresses = ""
            $forwardingTarget = ""
            $forwardingBehavior = ""
            $mailboxForwarded = $false
        }

        #--- Sign-in Status ---
        $signInStatus = if ($null -ne $user.AccountEnabled) { 
            if ($user.AccountEnabled) { "Enabled" } else { "Disabled" }
        } else { "Unknown" }

        #--- Combined Object ---
        $obj = [PSCustomObject]@{
            # Azure AD Properties
            GivenName                   = if ($null -ne $user.GivenName) { $user.GivenName } else { "" }
            Surname                     = if ($null -ne $user.Surname) { $user.Surname } else { "" }
            DisplayName                 = if ($null -ne $user.DisplayName) { $user.DisplayName } else { "" }
            UserPrincipalName           = if ($null -ne $user.UserPrincipalName) { $user.UserPrincipalName } else { "" }
            JobTitle                    = if ($null -ne $user.JobTitle) { $user.JobTitle } else { "" }
            Mail                        = if ($null -ne $user.Mail) { $user.Mail } else { "" }
            Department                  = if ($null -ne $user.Department) { $user.Department } else { "" }
            CompanyName                 = if ($null -ne $user.CompanyName) { $user.CompanyName } else { "" }
            OfficeLocation              = if ($null -ne $user.OfficeLocation) { $user.OfficeLocation } else { "" }
            EmployeeID                  = if ($null -ne $user.EmployeeId) { $user.EmployeeId } else { "" }
            MobilePhone                 = if ($null -ne $user.MobilePhone) { $user.MobilePhone } else { "" }
            BusinessPhones              = if ($null -ne $user.BusinessPhones) { ($user.BusinessPhones -join ';') } else { "" }
            StreetAddress               = if ($null -ne $user.StreetAddress) { $user.StreetAddress } else { "" }
            City                        = if ($null -ne $user.City) { $user.City } else { "" }
            PostalCode                  = if ($null -ne $user.PostalCode) { $user.PostalCode } else { "" }
            State                       = if ($null -ne $user.State) { $user.State } else { "" }
            Country                     = if ($null -ne $user.Country) { $user.Country } else { "" }
            UserType                    = if ($null -ne $user.UserType) { $user.UserType } else { "" }
            OnPremisesSyncEnabled       = if ($null -ne $user.OnPremisesSyncEnabled) { $user.OnPremisesSyncEnabled } else { "" }
            AccountEnabled              = if ($null -ne $user.AccountEnabled) { $user.AccountEnabled } else { "" }
            CreatedDateTime             = if ($null -ne $user.CreatedDateTime) { $user.CreatedDateTime.ToString("o") } else { "" }

            # Sign-in Status
            SignInStatus                = $signInStatus

            # License Information
            LicenseAssigned             = $licenseAssigned
            LicenseNames                = $licenseNames

            # Mailbox Properties
            PrimarySmtpAddress          = if ($mailbox -and $mailbox.PrimarySmtpAddress) { $mailbox.PrimarySmtpAddress } else { "" }
            AllEmailAddresses           = $emailAddresses
            RecipientTypeDetails        = if ($mailbox -and $mailbox.RecipientTypeDetails) { $mailbox.RecipientTypeDetails.ToString() } else { "" }
            AccountDisabled             = if ($mailbox -and $mailbox.AccountDisabled) { $mailbox.AccountDisabled } else { "" }
            HiddenFromAddressListsEnabled = if ($mailbox -and $mailbox.HiddenFromAddressListsEnabled) { $mailbox.HiddenFromAddressListsEnabled } else { "" }
            MailboxHiddenFromAddressLists = $mailboxHidden
            DeliverToMailboxAndForward  = if ($mailbox -and $mailbox.DeliverToMailboxAndForward) { $mailbox.DeliverToMailboxAndForward } else { "" }

            # Forwarding Details
            MailboxForwarded            = $mailboxForwarded
            ForwardingTarget            = $forwardingTarget
            ForwardingType              = $forwardingBehavior

            ForwardingAddress           = if ($mailbox -and $mailbox.ForwardingAddress) { $mailbox.ForwardingAddress } else { "" }
            ForwardingSmtpAddress       = if ($mailbox -and $mailbox.ForwardingSmtpAddress) { $mailbox.ForwardingSmtpAddress } else { "" }
            ArchiveStatus               = if ($mailbox -and $mailbox.ArchiveStatus) { $mailbox.ArchiveStatus } else { "" }

            # Mailbox Statistics
            ItemCount                   = if ($mailboxStats) { $mailboxStats.ItemCount } else { "" }
            TotalItemSize               = if ($mailboxStats) { $mailboxStats.TotalItemSize } else { "" }
            LastLogonTime               = if ($mailboxStats) { $mailboxStats.LastLogonTime } else { "" }

            # Archive Statistics
            ArchiveItemCount            = if ($archiveStats) { $archiveStats.ItemCount } else { "" }
            ArchiveTotalItemSize        = if ($archiveStats) { $archiveStats.TotalItemSize } else { "" }

            # OneDrive Properties
            OneDriveProvisioned         = if ($oneDriveProvisioned) { "Yes" } else { "No" }
            OneDriveUrl                 = $oneDriveUrl
            OneDriveLastActivity        = $lastActivity
            OneDriveFileCount           = $fileCount
            OneDriveUsedGB              = $usedGB
        }

        $combinedData += $obj
    }
    catch {
        Write-Warning "Error processing $($user.UserPrincipalName): $_"
        "[$(Get-Date -Format o)] $($user.UserPrincipalName): $_" | Out-File $ErrorLog -Append
    }
}

# Cleanup OneDrive report
Remove-Item $OneDriveReportPath -ErrorAction SilentlyContinue

# Export to CSV and cleanup connections
$combinedData | Export-Csv -Path $OutputFile -NoTypeInformation -Encoding UTF8
Disconnect-MgGraph
Disconnect-ExchangeOnline -Confirm:$false

Write-Output "Complete export saved to: $OutputFile"
if (Test-Path $ErrorLog) {
    Write-Warning "Some errors were logged. See $ErrorLog"
}
