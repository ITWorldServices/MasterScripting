$csv = Import-Csv -Path "PATH TO CSV"

ForEach ($item In $csv) 
    { 
        $create_group = New-ADGroup -Name $item.GroupName -GroupCategory $item.GroupCategory -groupScope $item.GroupScope -Path $item.OU 
        Write-Host -ForegroundColor Green "Group $($item.GroupName) created!" 
    }