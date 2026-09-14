# Import modules
$MaximumFunctionCount = 32768
Import-Module Microsoft.Graph
Import-Module ExchangeOnlineManagement

# Connect to Microsoft Graph and Exchange Online
Connect-MgGraph -Scopes "User.Read.All", "Directory.Read.All"
Connect-ExchangeOnline

# Get all users
$users = Get-MgUser -All

# Initialize an array to store user data
$userData = @()

foreach ($user in $users) {
    # Get Azure AD properties
    try {
        $azureADProps = Get-MgUser -UserId $user.Id -Property DisplayName, UserPrincipalName, GivenName, Surname, JobTitle, Department, OfficeLocation, MobilePhone, EmployeeId, AccountEnabled, CreatedDateTime, StreetAddress, City, State, PostalCode, Country, CompanyName, BusinessPhones, MailNickname

        # Get manager information
        $managerInfo = Get-MgUserManager -UserId $user.Id -ErrorAction SilentlyContinue
        if ($managerInfo) {
            $managerDisplayName = $managerInfo.AdditionalProperties.displayName
        } else {
            $managerDisplayName = "No Manager"
        }
    }
    catch {
        Write-Host "Error fetching Azure AD properties for user $($user.UserPrincipalName): $_" -ForegroundColor Red
        continue
    }

    # Initialize Exchange properties
    $exchangeProps = $null
    $mailboxStats = $null
    $archiveStats = $null
    $proxyAddresses = $null

    # Check if user has a mailbox
    try {
        $mailbox = Get-Mailbox -Identity $user.UserPrincipalName -ErrorAction Stop
        if ($mailbox) {
            # Get Exchange Online properties
            $exchangeProps = $mailbox | Select-Object PrimarySmtpAddress, RecipientTypeDetails, MailboxPlan, LitigationHoldEnabled, ArchiveStatus, DeliverToMailboxAndForward, ForwardingSmtpAddress
            
            # Get proxy addresses
            $proxyAddresses = $mailbox.EmailAddresses | Where-Object {$_ -clike "smtp:*"}
            
            # Get mailbox statistics
            $mailboxStats = Get-MailboxStatistics -Identity $user.UserPrincipalName
            
            # Get archive statistics if archive is enabled
            if ($mailbox.ArchiveStatus -eq 'Active') {
                $archiveStats = Get-MailboxStatistics -Identity $user.UserPrincipalName -Archive
            }
            
            Write-Host "User $($user.UserPrincipalName) has a mailbox." -ForegroundColor Green
        }
    }
    catch {
        Write-Host "User $($user.UserPrincipalName) does not have a mailbox." -ForegroundColor Yellow
    }

    # Combine properties
    $userObject = [PSCustomObject]@{
        # Azure AD properties
        DisplayName = $azureADProps.DisplayName
        UserPrincipalName = $azureADProps.UserPrincipalName
        MailNickName = $azureADProps.MailNickname
        GivenName = $azureADProps.GivenName
        Surname = $azureADProps.Surname
        JobTitle = $azureADProps.JobTitle
        Department = $azureADProps.Department
        OfficeLocation = $azureADProps.OfficeLocation
        MobilePhone = $azureADProps.MobilePhone
        TelephoneNumber = $azureADProps.BusinessPhones[0]
        EmployeeId = $azureADProps.EmployeeId
        AccountEnabled = $azureADProps.AccountEnabled
        CreatedDateTime = $azureADProps.CreatedDateTime
        StreetAddress = $azureADProps.StreetAddress
        City = $azureADProps.City
        State = $azureADProps.State
        PostalCode = $azureADProps.PostalCode
        Country = $azureADProps.Country
        CompanyName = $azureADProps.CompanyName
        ManagerDisplayName = $managerDisplayName
        
        # Exchange Online properties
        PrimarySmtpAddress = $exchangeProps.PrimarySmtpAddress
        RecipientTypeDetails = $exchangeProps.RecipientTypeDetails
        MailboxPlan = $exchangeProps.MailboxPlan
        LitigationHoldEnabled = $exchangeProps.LitigationHoldEnabled
        DeliverToMailboxAndForward = $exchangeProps.DeliverToMailboxAndForward
        ForwardingSmtpAddress = $exchangeProps.ForwardingSmtpAddress
        ProxyAddresses = ($proxyAddresses -join "; ")
        TotalItemSize = $mailboxStats.TotalItemSize
        ItemCount = $mailboxStats.ItemCount
        ArchiveStatus = $exchangeProps.ArchiveStatus
        ArchiveTotalItemSize = if ($archiveStats) { $archiveStats.TotalItemSize } else { "N/A" }
        ArchiveItemCount = if ($archiveStats) { $archiveStats.ItemCount } else { "N/A" }
    }
    
    $userData += $userObject
}

# Export to CSV
$userData | Export-Csv -Path "C:\temp\UserExport.csv" -NoTypeInformation

# Disconnect sessions
Disconnect-MgGraph
Disconnect-ExchangeOnline -Confirm:$false