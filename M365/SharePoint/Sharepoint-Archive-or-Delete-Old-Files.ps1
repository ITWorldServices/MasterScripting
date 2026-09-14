<#
.SYNOPSIS
    Cleans up old files in a SharePoint library by deleting or archiving.
.DESCRIPTION
    Uses PnP.PowerShell to find files not modified or accessed recently and deletes or archives them.

.HOW TO / USAGE EXAMPLES
    # 1. Run interactively (guided prompts for all parameters):
    .\Sharepoint-Archive-or-Delete-Old-Files.ps1

    # 2. Enable audit logging only (non-interactive):
    Enable-SharePointAuditLogging -SiteUrl 'https://contoso.sharepoint.com/sites/yourSite'

    # 3. Delete old files only (non-interactive):
    Delete-OldSharePointFiles -SiteUrl 'https://contoso.sharepoint.com/sites/yourSite' -LibraryName 'Documents' -DaysSinceModified 120 -DaysSinceAccessed 365

    # 4. Archive old files only (non-interactive):
    Archive-OldSharePointFiles -SiteUrl 'https://contoso.sharepoint.com/sites/yourSite' -LibraryName 'Documents' -DaysSinceModified 180 -DaysSinceAccessed 365 -ArchiveLibraryUrl '/sites/archive/Documents'

    # 5. Run with all parameters specified (no prompts):
    powershell.exe -File "C:\Path\To\Sharepoint-Archive-or-Delete-Old-Files.ps1" -SiteUrl "https://contoso.sharepoint.com/sites/yourSite" -LibraryName "Documents" -DaysSinceModified 90 -DaysSinceAccessed 180 -Action "Delete"

    # 6. Use in a scheduled task (Windows example):
    schtasks /Create /SC WEEKLY /TN "SPCleanup" /TR "powershell.exe -File C:\Path\To\Sharepoint-Archive-or-Delete-Old-Files.ps1 -SiteUrl 'https://contoso.sharepoint.com/sites/yourSite' -LibraryName 'Documents' -DaysSinceModified 90 -DaysSinceAccessed 180 -Action 'Archive' -ArchiveLibraryUrl '/sites/archive/Documents'" /ST 02:00

    # 7. Use in a script pipeline:
    $sites = @('https://contoso.sharepoint.com/sites/HR', 'https://contoso.sharepoint.com/sites/Finance')
    foreach ($site in $sites) {
        Delete-OldSharePointFiles -SiteUrl $site -LibraryName 'Documents' -DaysSinceModified 90 -DaysSinceAccessed 180
    }

.NOTES
    - Requires PnP.PowerShell module (Install-Module PnP.PowerShell -Force)
    - You must have appropriate permissions to connect and modify files in the target SharePoint site.
    - Audit logging must be enabled to use access-based filtering.
    - For archiving, specify a valid ArchiveLibraryUrl.
#>

# Requires PnP.PowerShell Module
# Ensure you install: Install-Module PnP.PowerShell -Force

function Enable-SharePointAuditLogging {
    param (
        [Parameter(Mandatory=$true)]
        [string]$SiteUrl
    )
    Write-Host "Connecting to SharePoint..." -ForegroundColor Cyan
    try {
        Connect-PnPOnline -Url $SiteUrl -Interactive
    } catch {
        Write-Error "Failed to connect to SharePoint: $_"
        exit 1
    }
    $auditSettings = Get-PnPAuditing
    if (-not $auditSettings.AuditFlags) {
        Write-Host "Enabling audit logging..." -ForegroundColor Yellow
        Set-PnPAuditing -EnableAll
        Write-Host "Audit logging enabled. Audit data will begin capturing from now forward." -ForegroundColor Green
    } else {
        Write-Host "Audit logging is already enabled." -ForegroundColor Green
    }
}

function Delete-OldSharePointFiles {
    param (
        [Parameter(Mandatory=$true)]
        [string]$SiteUrl,
        [Parameter(Mandatory=$true)]
        [string]$LibraryName,
        [Parameter(Mandatory=$false)]
        [int]$DaysSinceModified = 90,
        [Parameter(Mandatory=$false)]
        [int]$DaysSinceAccessed = 180
    )
    Write-Host "Connecting to SharePoint..." -ForegroundColor Cyan
    try {
        Connect-PnPOnline -Url $SiteUrl -Interactive
    } catch {
        Write-Error "Failed to connect to SharePoint: $_"
        exit 1
    }
    $modifiedThresholdDate = (Get-Date).AddDays(-$DaysSinceModified)
    $accessedThresholdDate = (Get-Date).AddDays(-$DaysSinceAccessed)
    Write-Host "Fetching items modified before $modifiedThresholdDate..." -ForegroundColor Cyan
    $items = Get-PnPListItem -List $LibraryName -PageSize 1000 | Where-Object { $_["Modified"] -lt $modifiedThresholdDate }
    $itemsToProcess = @()
    foreach ($item in $items) {
        $itemUrl = $item.FieldValues["FileRef"]
        $auditLogs = Get-PnPAuditLog -ContentType File -StartTime $accessedThresholdDate.AddYears(-1) -EndTime (Get-Date) -Url $itemUrl -ErrorAction SilentlyContinue
        if ($auditLogs) {
            $lastAccessed = ($auditLogs | Sort-Object -Property EventDate -Descending | Select-Object -First 1).EventDate
            if ($lastAccessed -lt $accessedThresholdDate) {
                $itemsToProcess += $item
            }
        } else {
            $itemsToProcess += $item
        }
    }
    Write-Host "Items to delete: $($itemsToProcess.Count)" -ForegroundColor Magenta
    $deletedCount = 0
    $skippedCount = 0
    foreach ($item in $itemsToProcess) {
        $fileUrl = $item.FieldValues["FileRef"]
        try {
            Remove-PnPListItem -List $LibraryName -Identity $item.Id -Force -Confirm:$false
            Write-Host "Deleted: $fileUrl" -ForegroundColor Green
            $deletedCount++
        } catch {
            Write-Warning "Error deleting $fileUrl : $_"
            $skippedCount++
        }
    }
    Write-Host "Delete process completed." -ForegroundColor Cyan
    Write-Host "Summary: Deleted=$deletedCount, Skipped=$skippedCount" -ForegroundColor Yellow
}

function Archive-OldSharePointFiles {
    param (
        [Parameter(Mandatory=$true)]
        [string]$SiteUrl,
        [Parameter(Mandatory=$true)]
        [string]$LibraryName,
        [Parameter(Mandatory=$false)]
        [int]$DaysSinceModified = 90,
        [Parameter(Mandatory=$false)]
        [int]$DaysSinceAccessed = 180,
        [Parameter(Mandatory=$true)]
        [string]$ArchiveLibraryUrl
    )
    Write-Host "Connecting to SharePoint..." -ForegroundColor Cyan
    try {
        Connect-PnPOnline -Url $SiteUrl -Interactive
    } catch {
        Write-Error "Failed to connect to SharePoint: $_"
        exit 1
    }
    $modifiedThresholdDate = (Get-Date).AddDays(-$DaysSinceModified)
    $accessedThresholdDate = (Get-Date).AddDays(-$DaysSinceAccessed)
    Write-Host "Fetching items modified before $modifiedThresholdDate..." -ForegroundColor Cyan
    $items = Get-PnPListItem -List $LibraryName -PageSize 1000 | Where-Object { $_["Modified"] -lt $modifiedThresholdDate }
    $itemsToProcess = @()
    foreach ($item in $items) {
        $itemUrl = $item.FieldValues["FileRef"]
        $auditLogs = Get-PnPAuditLog -ContentType File -StartTime $accessedThresholdDate.AddYears(-1) -EndTime (Get-Date) -Url $itemUrl -ErrorAction SilentlyContinue
        if ($auditLogs) {
            $lastAccessed = ($auditLogs | Sort-Object -Property EventDate -Descending | Select-Object -First 1).EventDate
            if ($lastAccessed -lt $accessedThresholdDate) {
                $itemsToProcess += $item
            }
        } else {
            $itemsToProcess += $item
        }
    }
    Write-Host "Items to archive: $($itemsToProcess.Count)" -ForegroundColor Magenta
    $archivedCount = 0
    $skippedCount = 0
    foreach ($item in $itemsToProcess) {
        $fileUrl = $item.FieldValues["FileRef"]
        try {
            $fileName = Split-Path $fileUrl -Leaf
            $destination = "$ArchiveLibraryUrl/$fileName"
            Move-PnPFile -SourceUrl $fileUrl -TargetUrl $destination -Force -Confirm:$false
            Write-Host "Archived: $fileUrl to $destination" -ForegroundColor Green
            $archivedCount++
        } catch {
            Write-Warning "Error archiving $fileUrl : $_"
            $skippedCount++
        }
    }
    Write-Host "Archive process completed." -ForegroundColor Cyan
    Write-Host "Summary: Archived=$archivedCount, Skipped=$skippedCount" -ForegroundColor Yellow
}

function Invoke-SharePointCleanupInteractive {
    Write-Host "=== SharePoint Library Cleanup ===" -ForegroundColor Cyan
    $action = Read-Host "Select action: EnableAuditLogging, Delete, Archive, or All (default: All)"
    if ([string]::IsNullOrEmpty($action)) { $action = 'All' }
    $SiteUrl = Read-Host "Enter SharePoint Site URL"
    if ($action -eq 'EnableAuditLogging') {
        Enable-SharePointAuditLogging -SiteUrl $SiteUrl
        return
    }
    $LibraryName = Read-Host "Enter Library Name"
    $DaysSinceModified = Read-Host "Enter number of days since last modification to qualify for processing [default=90]"
    if ([string]::IsNullOrEmpty($DaysSinceModified)) { $DaysSinceModified = 90 }
    $DaysSinceModified = [int]$DaysSinceModified
    $DaysSinceAccessed = Read-Host "Enter number of days since last accessed to qualify for processing [default=180]"
    if ([string]::IsNullOrEmpty($DaysSinceAccessed)) { $DaysSinceAccessed = 180 }
    $DaysSinceAccessed = [int]$DaysSinceAccessed
    if ($action -eq 'Delete') {
        Delete-OldSharePointFiles -SiteUrl $SiteUrl -LibraryName $LibraryName -DaysSinceModified $DaysSinceModified -DaysSinceAccessed $DaysSinceAccessed
        return
    }
    if ($action -eq 'Archive') {
        $ArchiveLibraryUrl = Read-Host "Enter Archive Library URL (e.g., /sites/archive/documents)"
        if ([string]::IsNullOrEmpty($ArchiveLibraryUrl)) {
            Write-Warning "No ArchiveLibraryUrl provided. Exiting."
            exit 1
        }
        Archive-OldSharePointFiles -SiteUrl $SiteUrl -LibraryName $LibraryName -DaysSinceModified $DaysSinceModified -DaysSinceAccessed $DaysSinceAccessed -ArchiveLibraryUrl $ArchiveLibraryUrl
        return
    }
    if ($action -eq 'All') {
        $EnableAuditLoggingPrompt = Read-Host "Enable audit logging if not already enabled? (Y/N) [default=N]"
        $EnableAuditLogging = $false
        if ($EnableAuditLoggingPrompt -eq 'Y') { $EnableAuditLogging = $true }
        if ($EnableAuditLogging) {
            Enable-SharePointAuditLogging -SiteUrl $SiteUrl
        }
        $DeleteOrArchive = Read-Host "Enter action for files ('Delete' or 'Archive') [default='Archive']"
        if ([string]::IsNullOrEmpty($DeleteOrArchive)) { $DeleteOrArchive = 'Archive' }
        if ($DeleteOrArchive -eq 'Delete') {
            Delete-OldSharePointFiles -SiteUrl $SiteUrl -LibraryName $LibraryName -DaysSinceModified $DaysSinceModified -DaysSinceAccessed $DaysSinceAccessed
        } else {
            $ArchiveLibraryUrl = Read-Host "Enter Archive Library URL (e.g., /sites/archive/documents)"
            if ([string]::IsNullOrEmpty($ArchiveLibraryUrl)) {
                Write-Warning "No ArchiveLibraryUrl provided. Exiting."
                exit 1
            }
            Archive-OldSharePointFiles -SiteUrl $SiteUrl -LibraryName $LibraryName -DaysSinceModified $DaysSinceModified -DaysSinceAccessed $DaysSinceAccessed -ArchiveLibraryUrl $ArchiveLibraryUrl
        }
    }
}

# Entry point for interactive use
if ($MyInvocation.InvocationName -eq '.') {
    # Script is dot-sourced, do not run interactively
    return
}
if ($PSCommandPath -eq $MyInvocation.MyCommand.Path) {
    Invoke-SharePointCleanupInteractive
}
