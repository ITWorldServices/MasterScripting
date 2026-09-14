# Define the list of folders to apply ACLs to
$folders = @(
    "Y:\ttbntfile1\shared\appsdata\groups\0-Policies and Procedures Manual",
    "Y:\ttbntfile1\shared\appsdata\groups\1-Administrative",
    "Y:\ttbntfile1\shared\appsdata\groups\10-Operations",
    "Y:\ttbntfile1\shared\appsdata\groups\11-TemplatesGeneralInfo",
    "Y:\ttbntfile1\shared\appsdata\groups\12-HR",
    "Y:\ttbntfile1\shared\appsdata\groups\13-Marketing",
    "Y:\ttbntfile1\shared\appsdata\groups\14-Accounting",
    "Y:\ttbntfile1\shared\appsdata\groups\15-Finance",
    "Y:\ttbntfile1\shared\appsdata\groups\16-DriversSupport",
    "Y:\ttbntfile1\shared\appsdata\groups\17-Xeroxscans",
    "Y:\ttbntfile1\shared\appsdata\groups\18-COIs",
    "Y:\ttbntfile1\shared\appsdata\groups\19-IT",
    "Y:\ttbntfile1\shared\appsdata\groups\2-Contracts",
    "Y:\ttbntfile1\shared\appsdata\groups\20-Executives",
    "Y:\ttbntfile1\shared\appsdata\groups\21-RealEstateServices",
    "Y:\ttbntfile1\shared\appsdata\groups\23-ModularizationDepartment",
    "Y:\ttbntfile1\shared\appsdata\groups\24-HospitalFacilitiesEngineering",
    "Y:\ttbntfile1\shared\appsdata\groups\25-Sub-ContractorAcknowledments",
    "Y:\ttbntfile1\shared\appsdata\groups\3-Design",
    "Y:\ttbntfile1\shared\appsdata\groups\4-PreconEstimating",
    "Y:\ttbntfile1\shared\appsdata\groups\5-Projects",
    "Y:\ttbntfile1\shared\appsdata\groups\6-ProgramMgmt",
    "Y:\ttbntfile1\shared\appsdata\groups\7-Quality",
    "Y:\ttbntfile1\shared\appsdata\groups\8-Safety",
    "Y:\ttbntfile1\shared\appsdata\groups\9-FieldLabor",
    "Y:\ttbntfile1\shared\appsdata\groups\Accounting",
    "Y:\ttbntfile1\shared\appsdata\groups\Exprpts",
    "Y:\ttbntfile1\shared\appsdata\groups\Finance",
    "Y:\ttbntfile1\shared\appsdata\groups\HPGL2",
    "Y:\ttbntfile1\shared\appsdata\groups\PM",
    "Y:\ttbntfile1\shared\appsdata\groups\Public",
    "Y:\ttbntfile1\shared\appsdata\groups\SC",
    "Y:\ttbntfile1\shared\appsdata\groups\StJames",
    "Y:\ttbntfile1\shared\appsdata\groups\Superintendent",
    "Y:\ttbntfile1\shared\appsdata\groups\Xerox Wide Format with FreeFlow Accxes Print Drivers 15.0.5 SIGNED",
    "Y:\ttbntfile1\shared\appsdata\groups\Xerox_Wide_Format_with_FreeFlow_Accxes",
    "Y:\ttbntfile1\shared\appsdata\groups\Xerox_Wide_Format_with_FreeFlow_Accxes_Print_Drivers"
)
 
 
# Define the ACL to apply (replace "YOUR_PERMISSIONS_HERE" with desired permissions)
$acl = '"tonnandblank\AVD Test Users":(OI)(CI)F'
 
 
# Iterate through folders and start a job for each folder
foreach ($folder in $folders) {
    Start-Job -ScriptBlock {
       param($folder, $acl)
        icacls $folder /inheritance:e
    } -ArgumentList $folder, $acl 
}
 
# Wait for all jobs to complete
Get-Job | Wait-Job
 
# Cleanup - remove completed jobs
Get-Job | Remove-Job