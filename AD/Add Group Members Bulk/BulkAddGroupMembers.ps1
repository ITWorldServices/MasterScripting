# Script for bulk add users to ad groups
#Import csv file and loads adusers in variable

$Users = Import-Csv -Path "PATH TO CSV FILE"

#Iterate AdUsers to add user account in Group

foreach($User in $Users){
        try
        {

            Add-ADGroupMember -Identity "ADGroupName" -Members $User.SamAccountName -ErrorAction Stop -Verbose
        }
        catch
        {
            Write-Host "Error while adding "$User.SamAccountName" to adgroup"
        }

    }