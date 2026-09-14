Get-AzureADUser -All $true | ForEach-Object {
    $manager = Get-AzureADUserManager -ObjectId $_.ObjectId
    [PSCustomObject]@{
        DisplayName = $_.DisplayName
        UserPrincipalName = $_.UserPrincipalName
        Department = $_.Department
        ManagerName = $manager.DisplayName
    }
} | Export-Csv -Path "C:\Temp\NewUsers.csv" -NoTypeInformation

