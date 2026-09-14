Importing Groups to Active Directory from ExchangeOnline Export or AD Export.

STEP 0:
Connect-ExchangeOnline - Powershell and enter in Global Admin creds

STEP 1: Export Groups

Get-UnifiedGroup | Export-Csv "PATH TO CSV SAVE LOCATION"

STEP 2: Editing the columns to match example

NOTE: In order to get a proper SamAccountName column, 

2a. First step:

you can copy and paste your PrimarySmtpAddress (you now have two identical columns) Rename one as SamAccountName - 
NOTE: Be sure your CSV doesn't already have a SamAccountName column, it is usually wrong format so you will need to delete the other one if it exists

2a. Second step:

then use the find & Select tool in Excel to remove the @*.* from each row, by targeting the single column you previously copied and pasted

MODIFY EMAIL ADDRESSES TO ONLY WHAT YOU NEED USING THE FIND & REPLACE TOOL

2b. modify the EmailAddress column to remove the unwanted details, this column has each email address separated by a SPACE, 

2b. First step:

Find field "empty space bar single" - Replace field "," without the quotes on both fields  

2b. Second step:

Further manipulate the data by removing the "SPO:SPO*@*," without the quotes, replace field should be empty, this will remove all the SPO garbage addresses

2b. Third step: 

Further remove the onmicrosoft.com email addresses as they are not needed. 

Entries at end or middle of email addresses column:

Find field ",smtp:*@ncslogistics.onmicrosoft.com" - Replace field "empty" - no quotes

Entries at beginning of email address column:

Find field "smtp:*@ncslogistics.onmicrosoft.com" - Replace field "empty" - no quotes

Result should be only the primary email addresses SMTP:blah@blah.com and any valid aliases smtp:blah@blah.com with a comma remaining and separating all of them - This comma will be needed for using the AddProxyAddress script

STEP 3: Create GroupCategory & GroupScope columns

Create two new columns titled as above

Category enter in category name for AD attribute (See example file)

Scope enter in the scope name for AD attribute (See example file)

STEP 4: Add in the destination OU path

See example file

To get the path, create the OU in active directory, make a test group, use powershell to find the OU path to copy to ensure you don't mess up

Get-ADGroup "testgroupname" without the quotes will give you the OU destination

NOTE: remove the groupname from this and that will give you the proper path

STEP 5: Remove all remaining unwanted columns and save the file as a "csv"


