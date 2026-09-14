Get-MgSubscribedSku | ForEach-Object {
    Write-Host "Product: $($_.SkuPartNumber)"
    $_.ServicePlans | ForEach-Object {
        Write-Host ("    " + $_.ServicePlanName + " : " + $_.ServicePlanId)
    }
}
