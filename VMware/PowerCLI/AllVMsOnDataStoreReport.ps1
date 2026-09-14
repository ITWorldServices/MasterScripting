# Ensure PowerCLI is installed
if (-not (Get-Module -ListAvailable -Name VMware.PowerCLI)) {
    Write-Host "PowerCLI not found. Installing VMware.PowerCLI module from PowerShell Gallery..."
    Install-Module -Name VMware.PowerCLI -Scope CurrentUser -Force
}
Import-Module VMware.PowerCLI

# Connect to vCenter Server (if not already connected)
Connect-VIServer -Server "vcenter.enerdomain.net" -User 'administrator@vsphere.local' -Password 'ChangeMe'

# Initialize an array to store report data
$report = @()

# Retrieve all VMs and their datastore usage
Get-VM | ForEach-Object {
    $vm = $_
    $vmName = $vm.Name

    # Iterate through each datastore that the VM is using
    $vm.ExtensionData.Storage.PerDatastoreUsage | ForEach-Object {
        $datastore = $_.Datastore.Name
        $provisionedSizeGB = [math]::Round($_.Committed / 1GB, 2)  # Provisioned (or allocated) size in GB
        $usedSizeGB = [math]::Round(($_.Committed + $_.Uncommitted) / 1GB, 2)  # Used size in GB

        # Add to report
        $report += [PSCustomObject]@{
            VMName           = $vmName
            Datastore        = $datastore
            ProvisionedSizeGB = $provisionedSizeGB
            UsedSizeGB       = $usedSizeGB
        }
    }
}

# Output the report
$report | Format-Table -AutoSize

# Optionally export to CSV
$report | Export-Csv -Path "C:\AllVMsDatastoresReport.csv" -NoTypeInformation -UseCulture

# Disconnect from vCenter (optional)
Disconnect-VIServer -Confirm:$false