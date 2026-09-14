
# Get the ObjectGUID from on-prem Active Directory for the user
# ONPREM DC ONPREM DC ONPREM DC #
# $guid = (Get-ADUser -Identity "user").ObjectGUID
# $immutableId = [System.Convert]::ToBase64String($guid.ToByteArray())

# 365 TENANT 365 TENANT 365 TENANT #
Connect-MgGraph -Scopes User.ReadWrite.All
Update-MgUser -UserId "katy.walsh@primocenter.org" -OnPremisesImmutableId $immutableId

