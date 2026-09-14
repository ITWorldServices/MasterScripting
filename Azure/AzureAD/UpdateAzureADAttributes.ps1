# Read user details from the CSV file
$AzureADUsers = Import-CSV "C:\Temp\TonnBlank\UpdateAttributes.csv"
$i = 0;
$TotalRows = $AzureADUsers.Count
 
# Array to add update status
$UpdateStatusResult=@()
 
# Iterate and set user details one by one
ForEach($UserInfo in $AzureADUsers)
{
$UserId = $UserInfo.'UserPrincipalName'
 
# Convert CSV user info (PSObject) to hashtable
$NewUserData = @{}
$UserInfo.PSObject.Properties | ForEach { $NewUserData[$_.Name] = $_.Value }
 
$i++;
Write-Progress -activity "Processing $UserId " -status "$i out of $TotalRows completed"
 
Try
{
 
# Get current Azure AD user object
$UserObj = Get-AzureADUser -ObjectId $UserId
 
# Convert current Azure AD user object to hashtable
$ExistingUserData = @{}
$UserObj.PSObject.Properties | ForEach { $ExistingUserData[$_.Name] = $_.Value }
 
$AttributesToUpdate = @{}
 
# The CSV header names should have the same member property name supported in the Get-AzureADUser cmdlet. 
# Run this command to get the supported properties: Get-AzureADUser | Get-Member -MemberType property
$CSVHeaders = @("JobTitle","Department","surname","givenName","city")
 
ForEach($property in $CSVHeaders)
{
# Check the CSV field has value and compare the value with existing user property value.
if ($NewUserData[$property] -ne $null -and ($NewUserData[$property] -ne $ExistingUserData[$property]))
{
$AttributesToUpdate[$property] = $NewUserData[$property]
}
}
if($AttributesToUpdate.Count -gt 0)
{
# Set required user attributes.
# Need to prefix the variable AttributesToUpdate with @ symbol instead of $ to pass hashtable as parameters (ex: @AttributesToUpdate).
Set-AzureADUser -ObjectId $UserId @AttributesToUpdate
$UpdateStatus = "Success - Updated attributes : " + ($AttributesToUpdate.Keys -join ',')
 
} else {
$UpdateStatus ="No changes required"
}
 
}
catch
{
$UpdateStatus = "Failed: $_"
}
 
# Add user update status
$UpdateStatusResult += New-Object PSObject -property $([ordered]@{
User = $UserId
Status = $UpdateStatus
})
}
 
# Display the user update status result
$UpdateStatusResult | Select User,Status
 
# Export the update status report to CSV file
#$UpdateStatusResult | Export-CSV "C:\AzureADUserUpdateStatus.CSV" -NoTypeInformation -Encoding UTF8