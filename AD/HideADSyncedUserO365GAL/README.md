# Hide Active Directory Synced User from Office 365 Global Address List
## Step 1: Scope in the msDS-cloudExtensionAttribute for Azure AD Connect
### Open the **Azure AD Connect Synchronization Service**
![Azure AD Connect Synchronization Service](https://github.com/Impact-Technical-Project-Team/PowerShell-Scripts/blob/8b5bb9ceb2cc17630c7159da1fe9c4ece8f83028/M365/Hide%20AD%20Synced%20User%20from%20O365%20GAL/images/Azure-AD-Connect-Synchronization-Service.png)
### Navigate to the **Connectors** tab, select your **Active Directory** (not the domain.onmicrosoft.com entry), and select **Properties**
![Azure AD Connect Synchronization Properties](https://github.com/Impact-Technical-Project-Team/PowerShell-Scripts/blob/4d6bc00108d396e1eb8121458e296a8c4d6ab704/M365/Hide%20AD%20Synced%20User%20from%20O365%20GAL/images/Azure-AD-Connect-Synchronization-Service-Manager-Connectors-Properties.png)
### In the top right, click on Show All, scroll down and find msDS-CloudExtensionAttribute1 (you can use any of the numbers 1-20, just make sure to check the box you are using), and select OK
![Azure AD Connect Synchronization Properties](https://github.com/Impact-Technical-Project-Team/PowerShell-Scripts/blob/4d6bc00108d396e1eb8121458e296a8c4d6ab704/M365/Hide%20AD%20Synced%20User%20from%20O365%20GAL/images/Azure-AD-Connect-Synchronization-Service-Manager-Connectors-Properties-msDS-cloudExtensionAttribute1.png)
## Step 2: Create a custom sync rule
### Open up the Azure AD Connect Synchronization Rules Editor
![Azure AD Connect Synchronization Service](https://github.com/Impact-Technical-Project-Team/PowerShell-Scripts/blob/8b5bb9ceb2cc17630c7159da1fe9c4ece8f83028/M365/Hide%20AD%20Synced%20User%20from%20O365%20GAL/images/Azure-AD-Connect-Synchronization-Service.png)
### Click on the Add new rule button (make sure direction in the top left shows Inbound)
![Azure AD Rule Editor - Add New Rule](https://github.com/Impact-Technical-Project-Team/PowerShell-Scripts/blob/5d5db845e2ac6f289d37dad5bf1bc9569e119e33/M365/Hide%20AD%20Synced%20User%20from%20O365%20GAL/images/Azure-AD-Connect-Synchronization-Rules-Editor-Add-new-rule.png)
### Enter the following for the description:
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Name: **Hide user from GAL**  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Description: **If msDS-CloudExtensionAttribute1 attribute is set to HideFromGAL, hide from Exchange Online GAL**  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Connected System: **Your Active Directory Domain Name**  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Connected System Object Type: **user**  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Metaverse Object Type: **person**  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Link Type: **Join**  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Precedence: **50** (this can be any number less than 100.  Just make sure you don't duplicate numbers if you have other custom rules or you'll receive a dead-lock error from SQL Server)  
![Azure AD Rule Editor - Add New Rule Description](https://github.com/Impact-Technical-Project-Team/PowerShell-Scripts/blob/14d1ea427d30b6492e14732a3be9af64f81d752e/M365/Hide%20AD%20Synced%20User%20from%20O365%20GAL/images/Azure-AD-Connect-Synchronization-Rules-Editor-Add-new-rule-Description.png)
Click **Next** > on **Scoping filter** and **Join rules**, those can remain blank

Enter the following Transformation page, click the **Add transformation** button, fill out the form with the values below, and then click **Add**  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;FlowType: **Expression**  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Target Attribute: **msExchHideFromAddressLists**  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Source:  
**```IIF(IsPresent([msDS-cloudExtensionAttribute1]),IIF([msDS-cloudExtensionAttribute1]="HideFromGAL",True,False),NULL)```**  
![Azure AD Rule Editor - Add New Rule Filter](https://github.com/Impact-Technical-Project-Team/PowerShell-Scripts/blob/14d1ea427d30b6492e14732a3be9af64f81d752e/M365/Hide%20AD%20Synced%20User%20from%20O365%20GAL/images/Azure-AD-Connect-Synchronization-Rules-Editor-Add-new-rule-Transformations-msExchHideFromAddressLists.png)
## Step 3: Perform an initial sync
### Open up Windows PowerShell on the Azure AD Connect Server
![ ](https://github.com/Impact-Technical-Project-Team/PowerShell-Scripts/blob/14d1ea427d30b6492e14732a3be9af64f81d752e/M365/Hide%20AD%20Synced%20User%20from%20O365%20GAL/images/Windows-Server-2016-Windows-PowerShell-Run-as-administrator.png)
Execute the following command: **Start-ADSyncSyncCycle -PolicyType Initial**
![ ](https://github.com/Impact-Technical-Project-Team/PowerShell-Scripts/blob/14d1ea427d30b6492e14732a3be9af64f81d752e/M365/Hide%20AD%20Synced%20User%20from%20O365%20GAL/images/Windows-PowerShell-Azure-AD-Connect-Start-ADSyncSyncCycle-PolicyType-Initial-Delta.png)
## Step 4: Hide a user from Active Directory
### Open Active Directory Users and Computers, find the user you want to hide from the GAL, right click select Properties
![ ](https://github.com/Impact-Technical-Project-Team/PowerShell-Scripts/blob/14d1ea427d30b6492e14732a3be9af64f81d752e/M365/Hide%20AD%20Synced%20User%20from%20O365%20GAL/images/Active-Directory-Users-and-Computers-Properties.png)
### Select the Attributes Editor tab, find msDS-cloudExtensionAttribute1, and enter the value HideFromGAL (note, this is case sensitive), click OK and OK to close out of the editor. 
![ ](https://github.com/Impact-Technical-Project-Team/PowerShell-Scripts/blob/14d1ea427d30b6492e14732a3be9af64f81d752e/M365/Hide%20AD%20Synced%20User%20from%20O365%20GAL/images/Active-Directory-Users-and-Computers-Properties-User-msDS-cloudExtensionAttribute1-HideFromGAL.png)  
Note: if you don't see the Attribute Editor tab in the previous step, within Active Directory Users and Computers, click on View in the top menu and select Advanced Features  
![ ](https://github.com/Impact-Technical-Project-Team/PowerShell-Scripts/blob/14d1ea427d30b6492e14732a3be9af64f81d752e/M365/Hide%20AD%20Synced%20User%20from%20O365%20GAL/images/Active-Directory-Users-and-Computers-View-Advanced-Features.png)
## Step 5: Validation
### Open the Azure AD Connect Synchronization Service
![ ](https://github.com/Impact-Technical-Project-Team/PowerShell-Scripts/blob/14d1ea427d30b6492e14732a3be9af64f81d752e/M365/Hide%20AD%20Synced%20User%20from%20O365%20GAL/images/Azure-AD-Connect-Synchronization-Service.png)
On the Operations tab, if you haven't seen a Delta Synchronization, manually trigger the Delta sync to pick up the change you made in Active Directory
![ ](https://github.com/Impact-Technical-Project-Team/PowerShell-Scripts/blob/14d1ea427d30b6492e14732a3be9af64f81d752e/M365/Hide%20AD%20Synced%20User%20from%20O365%20GAL/images/Windows-PowerShell-Azure-AD-Connect-Start-ADSyncSyncCycle-PolicyType-Initial-Delta.png)
### Select the Export for the domain.onmicrosoft.com connecter and you should see 1 Updates
![ ](https://github.com/Impact-Technical-Project-Team/PowerShell-Scripts/blob/14d1ea427d30b6492e14732a3be9af64f81d752e/M365/Hide%20AD%20Synced%20User%20from%20O365%20GAL/images/Azure-AD-Connect-Synchronization-Service-Manager-Export-Updates-1.png)
### Select the user account that is listed and click Properties.  On the Connector Space Object Properties, you should see Azure AD Connect triggered an add to Azure AD to set msExchHideFromAddressLists set to true
![ ](https://github.com/Impact-Technical-Project-Team/PowerShell-Scripts/blob/14d1ea427d30b6492e14732a3be9af64f81d752e/M365/Hide%20AD%20Synced%20User%20from%20O365%20GAL/images/Azure-AD-Connect-Synchronization-Service-Manager-Properties-Connector-Space-Object-Properties-msExchHideFromAddressLists.png)
There ya have it!  An easy way to hide users from the GAL with minimal risk to ongoing operations.  Due to the way Azure AD Connect upgrades, our sync rule will persist fine during regular updates/patches released.