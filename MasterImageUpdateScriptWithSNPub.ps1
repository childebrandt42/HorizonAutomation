# Horizon Master Image Automated Update Process
# Chris Hildebrandt
# 12-5-2018
# Version 5.1 updated 2-24-2019
# Script will power on all master, Run software updates, and run the logoff Script, Snapshot the Masters, Update the Notes, Recompose or Push Images
# Change Notes: Removed the dependecy of user input vCenter info, and Folder use, added ServiceNow functions, and fixed some minor reporting issues.

#______________________________________________________________________________________
#User Editable Varibles

$HVServers = @("ServerName.Domain.com","ServerName.Domain.com")   #Enter HVServer FQDN in format of ("HVServer1","HVServer2")   
$SnapDays = "-30"   #Number of Days to keep past snapshots on the VDI Master VM must be a negitive number
$RCTimeDelay1 = '1' #Recompose Test Pools Delay in hours
$RCTimeDelay2 = '48'    #Recompose Prod Pools Delay in hours
$TestPoolNC = "Test"    #Test pool naming convention that diffirentiates them from standard pool.

$LogLocation = "C:\VDI_Tools\Logs"  #Location to save the logs on the Master VM. They will be copied to share and deleted upon completion of the script.
$ScriptLocation = "C:\VDI_Tools\Scripts"    #Location for the Scripts on the Master VMs. This folder will remain after script completes
$CloneToolsLocation = "C:\VDI_Tools\CloneTools" #Location on the Master VMs where the VMware Optimizer is stored and any other tools.
$SleepTimeInS = "180"   #Wait timer for windows to install Windows Updates via Windows Update Service. NOT needed if using SCCM to install updates. 
$AdobeUpdate = "AdobeAcrobatUpdate.ps1" #Name of Adobe Update Script. NOT needed if using SCCM to install updates.
$FlashUpdate = "FlashUpdate.ps1"    #Name of Flash Update Script. NOT needed if using SCCM to install updates.
$FirefoxUpdate = "FireFoxUpdate.ps1"    #Name of FireFox Update Script. NOT needed if using SCCM to install updates.
$JavaUpdate = "JavaUpdate.ps1"  #Name of Java Update Script. NOT needed if using SCCM to install updates.
$CustomRegPath = "HKLM:\Software\YourCompanyName\Horizon\" #Custom Registry Path to create a Key that keeps record of when it was last updated.
$RunOptimizer = '1' #To run Optimizer enter 1, If you do not want to run enter 0
$RunSEP = '1' #To run Optimizer enter 1, If you do not want to run enter 0
$RunServiceNow = '1'    #To Create Service Now Incidents and Change Tickets with details of the Recompose process. 
$OptimizerTemplateNamingPrefix = "TemplateName" #Template Naming i.e. "CompanyTemplate10" for windows 10 Please make sure to append the namy by 10 for win10, 8 for win8, and 7 for win7
$VMwareOptimizerName = "VMwareOSOptimizationTool.exe"   #Name Of the VMware Optimizer Tool
$ShareLogLocation = "\\ShareName.domain.com\VDI_Tools\Logs" #Network Share to save the logs. This will be the collection point for all logs as each VDI Master VM will copy there local logs to this location.
$ShareScriptLocation = "\\ShareName.domain.com\VDI_Tools\Scripts"   #Network Share location where all the scripts are stored and distributed from. 
$ShareCloneToolsLocation = "\\ShareName.domain.com\VDI_Tools\CloneTools"    #Network Share Location where all the Clone Tools are stored and distributed from.
$DomainName = "Domain.Name"     #The Domain that the Master Images are Joined to.
$CorporateBuildRegistryKeyPath = "HKLM:\Software\YourCompanyName\OS\"
$CorporateBuildRegistryKeyName = "BuildVersion"

#______________________________________________________________________________________
# Custom Registry, VMware Notes, and Attributes names. 
$RegNameCorporateBuildVersion = "Corporate Build Version" 
$RegNameWindowsVersion = "Windows Version"
$RegNameWindowsBuildNumber = "Windows Build Number"
$RegNameWindowsRevision = "Windows Revision"
$RegNameLastUpdateDate = "Last Update Date"
$RegNameLastRecomposeDate = "Last Recompose Date" 
$RegNamePoolType =  "Pool Type"
$RegNamePoolProvisionType = "Pool Provision Type"
$RegNamePoolAssignmentType = "Pool Assignment Type"
$RegNamePoolName = "Pool Name"

#______________________________________________________________________________________
#ServiceNow Varibles
$SNAddress = "https://YOURSERVICENOW.service-now.com"

#______________________________________________________________________________________
#Incident Varribles
$SNINCCallerID = "Caller Sys_ID" #Look up the Caller User you want to uses Sys_ID in your Service Now instance
$SNINCUrgency = "2" #Look this up in your Service Now instance
$SNINCImpact = "3" #Look this up in your Service Now instance
$SNINCPriority = "4" #Look this up in your Service Now instance
$SNINCContactType = "email" #Look this up in your Service Now instance
$SNINCNotify = "2" #Look this up in your Service Now instance
$SNINCWatchlist = "Watch List Sys_ID" #Can do comma seperated users Sys_ID's
$SNINCServiceOffering = "Service Offering" #Look this up in your Service Now instance
$SNINCProductionImpact = "No" #Well I hope its a No. 
$SNINCCategory = "Your Catagory" #Look this up in your Service Now instance
$SNINCSubcategory = "Your SubCat" #Look this up in your Service Now instance 
$SNINCItem = "request" #Look this up in your Item menu
$SNINCAssignmentGroup = "Assignment Group Sys_ID" #Look up the Assignment group you want to uses Sys_ID in your Service Now instance
$SNINCAssignedTo = "Assigned To Sys_ID" #Look up the Assignment to user you want to uses Sys_ID in your Service Now instance
$SNINCShortDescription = "Weekly Patch Cycle Updates VDI Pools for $SNDate" #Short Descriptiong of the task
$SNINCDescription = "Weekly Patch Cycle Update for VDI pools.
The following VDI Master Images are being updated with the latest Windows updates and 3rd Party software.
$HVPoolMSTR
Test Pools will be Refreshed at $SNScriptDate
Production Pools will be Refreshed $SNScriptDate2" #Full description of what you are trying to acomplish this is my example.

#______________________________________________________________________________________
#Change Varribles
$SNCHGRequestedBy = "Requested By Sys_ID" #Look up the Requested by User you want to uses Sys_ID in your Service Now instance, Normaly the same as the Caller for the INC
$SNCHGCategory = "Change Catagory"  #Look this up in your Service Now instance
$SNCHGServiceOffering = "Service Offering"  #Look this up in your Service Now instance
$SNCHGReason = "Change Reason" #Look this up in your Service Now instance
$SNCHGClientImpact = "No" #Look this up in your Service Now instance
$SNCHGStartDate = (Get-Date).ToString('yyyy-MM-dd HH:mm:ss') #date in string format. Only way it works.
$SNCHGEndDate = (Get-Date).AddHours($RCTimeDelay2+24).ToString('yyyy-MM-dd HH:mm:ss') #24 hour delay of right now (Can Change if needed) Has to be as a string
$SNCHGWatchList = "Watch List Sys_ID" #Can do comma seperated users Sys_ID's
$SNCHGUrgency = "2" #Look this up in your Service Now instance
$SNCHGRisk = "4" #Look this up in your Service Now instance
$SNCHGType = "Standard" #Look this up in your Service Now instance
$SNCHGState = "1" #Look this up in your Service Now instance
$SNCHGAssignmentGroup = "Assignment Group Sys_ID" #Look up the Assignment Group you want to uses Sys_ID in your Service Now instance
$SNCHGAssignedTo = "Assigned To Sys_ID" #Look up the Assigned to User you want to uses Sys_ID in your Service Now instance

#______________________________________________________________________________________
# Example of what I am using for my change information. 
$SNCHGJustification = "Weekly Sercurity Patching to install latest Windows updates, 3rd Party updates and Symantec antivirus client update"
$SNCHGChangePlan = "Run report on each VDI Master Image for install Software
Install Windows Updates from SCCM
Install 3rd Party software updates from SCCM
Run Report on each VDI Master Image for installed Software after updates have been completed
Update Custom Reg key to reflect the last date updated
Reboot VDI master VM
Run Shutdown Script that will Run Disk Cleanup, Defragment C drive, Pre-Compile .NET Framework, Run SEP Update, Run full system Scan, Force check-in with SEP server, Run VMware Optimization Tool, Clean out DownLoad's Cache Folder, Clear Event Logs, Release IP, Clear DNS, and Shutdown the VM.
Create a vCenter Snapshot of each VDI Master Image.
Update the vCenter VDI Master Image notes to show current recompose date
Remove Old Snapshots from the VDI Master Images
Power Back on the VDI Master Image
Start all the services
Copy Logs from VDI Master Images to remote Share and upload to this Change.
Refresh each of the pools based on Prod and Test timelines."
$SNCHGTestPlan = "All VDI Clone Pools listed as Test Pools will Be updated 1 hour after script completion, and Production pools will be updated 48 hours later."
$SNCHGBackoutPlan = "If there is a error found, will cancel all future Refresh tasks, and revert the ones that have been refreshed already to the previous snapshot."

#______________________________________________________________________________________
# Do not Edit Below! Do not Edit Below! Do not Edit Below! Do not Edit Below!
#______________________________________________________________________________________

#______________________________________________________________________________________
#Enviroment Varribles Do not EDIT
$ScriptDate = Get-Date -Format MM-dd-yyyy
$SNDate = (get-date).ToString('MM-dd-yyyy')
$SNScriptDate = (Get-Date).AddHours($RCTimeDelay1)
$SNScriptDate2 = (Get-Date).AddHours($RCTimeDelay2)

#______________________________________________________________________________________
#ServiceNow Method Varibles Do not edit these
$SNMethodPost = "post"
$SNMethodGet = "get"
$SNMethodPut = "put"
$SNMethodPatch = "patch"
$SNINCAddress = "$SNAddress/api/now/table/incident"
$SNCHGAddress = "$SNAddress/api/now/table/change_request"

#______________________________________________________________________________________
#Start the Debug Logging
$ErrorActionPreference="SilentlyContinue"
Stop-Transcript | out-null
$ErrorActionPreference = "Continue"
Start-Transcript -Force -Path $LogLocation\$ScriptDate\WeeklyVDIUpdates.txt
Write-Host "Start Debug Logs"

#______________________________________________________________________________________
#Import VMware Modules
#Install-Module -Name VMware.PowerCLI
#Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Force -confirm:$false
Get-Module -ListAvailable -Name VMware* | Import-Module
Set-PowerCLIConfiguration -Scope User -ParticipateInCEIP $false -confirm:$false
Set-PowerCLIConfiguration -InvalidCertificateAction ignore -confirm:$false
Write-Host "Enable all VMware Modules out of lazyness"


#______________________________________________________________________________________
#Load Functions
#______________________________________________________________________________________

#______________________________________________________________________________________
#Import Start Services Function
function Start-VDIservices ($SVCname) {
    Set-Service -Name $svcname -StartupType Automatic
    write-host "Changed Service startyp type for $Svcname to the following"
    Get-Service -Name $svcname | Select-Object StartType
    Start-Service -Name $svcname
    Write-Host "$SVCname running status is:"
    Get-Service -Name $svcname | Select-Object Status
    Do {
    $svc = Get-Service -Name $svcname
    Start-Sleep 2
    Write-Host "$SVCname running status is:"
    $svc
    } While ( $svc.Status -ne "Running" )
    }

#______________________________________________________________________________________
#Iport Function To Create Folder Path on VDI Master Images
Function Create-Path ($TestPath) 
    {
    if(!(Test-Path -Path $TestPath))
    {
    New-Item -Path "$TestPath" -ItemType "directory"
    Write-host "Created Directory $TestPath on Master Image $VMline" 
    }
    }

#______________________________________________________________________________________
#Import Function to Check to see if 3rd Party Script exists on share and copy to VDI Master Images
Function If-TestPath ($ShareTestPath)
    {
    if(Test-Path -Path "$ShareScriptLocation\$ShareTestPath")
    {
    Copy-Item -Path $ShareScriptLocation\$ShareTestPath -Destination $MasterScriptLocation
    Write-Host "Copy Script $ShareTestPath to folder $MasterScriptLocation on Master Image $VMLine"
    }
    }

#______________________________________________________________________________________
#Import Passwords
#______________________________________________________________________________________

#______________________________________________________________________________________
#Check is VMware Service Account Password file exists
if(-Not (Test-Path -Path "$ScriptLocation\VMwareServiceAccntPassword.txt" ))
{
    #______________________________________________________________________________________
    #Create Secure Password File
    Get-Credential -Message "Enter VMware Service Account Domain\Username" | Export-Clixml "$ScriptLocation\VMwareServiceAccntPassword.txt"
    Write-Host "Created Secure Credentials for vCenter and Horizon from Text file"
}

#______________________________________________________________________________________
#Import Secure Creds for use.
$VMwareSVCCreds = Import-Clixml "$ScriptLocation\VMwareServiceAccntPassword.txt"
Write-Host "Imported Secure Credentials for vCenter and Horizon from Text file"

#______________________________________________________________________________________
#Import Service Now Creditial for Service Now, and Define Headers and Import Functions
if($RunServiceNow -eq '1')
{
#______________________________________________________________________________________
#Create Service Now Creds if they do not already exist
if(-Not (Test-Path -Path "C:\VDI_Tools\Scripts\SNProdAccount.txt" ))
{
    Get-Credential -Message "Enter Your ServiceNow Account! Username@domain" | Export-Clixml "C:\VDI_Tools\Scripts\SNProdAccount.txt"
    Write-Host "Created Secure Credentials for ServiceNow"
}

#______________________________________________________________________________________
#Import Service Now Creds
$SNCreds = Import-Clixml "C:\VDI_Tools\Scripts\SNProdAccount.txt"
Write-Host "Imported Secure Credentials for ServiceNow from Text file"

#______________________________________________________________________________________
#Decrypt Password to imput into ServiceNow
$SNTextPass = $SNCreds.Password | ConvertFrom-SecureString
$SNTextPassPlain = [Runtime.InteropServices.Marshal]::PtrToStringAuto([Runtime.InteropServices.Marshal]::SecureStringToBSTR( (ConvertTo-SecureString $SNTextPass) ))
Write-host "Decrypt Master Password to clear text for for use for ServiceNow"

#______________________________________________________________________________________
#Create ServiceNow Creds Varribles
$SNuser = $SNCreds.Username
$SNpass = $SNTextPassPlain

#______________________________________________________________________________________
#ServiceNow Method Varibles Do not edit these
$SNMethodPost = "post"
$SNMethodGet = "get"
$SNMethodPut = "put"
$SNMethodPatch = "patch"

$SNINCAddress = "$SNAddress/api/now/table/incident"
$SNCHGAddress = "$SNAddress/api/now/table/change_request"

#______________________________________________________________________________________
#ServiceNow Build Auth Headers
$base64AuthInfo = [Convert]::ToBase64String([Text.Encoding]::ASCII.GetBytes(("{0}:{1}" -f $SNuser, $SNpass)))

#______________________________________________________________________________________
#ServiceNow Set Header
$SNheaders = New-Object "System.Collections.Generic.Dictionary[[String],[String]]"
$SNheaders.Add('Authorization',('Basic {0}' -f $base64AuthInfo))
$SNheaders.Add('Accept','application/json')
$SNheaders.Add('Content-Type','application/json')

#______________________________________________________________________________________
#Import Create-Incident Function
Function Create-Incident()
{
#______________________________________________________________________________________
# Specify request body
$SNCreateINCBody = @{ #Create Body of the Post Request
    caller_id= $SNINCCallerID
    urgency= $SNINCUrgency
    impact= $SNINCImpact
    priority= $SNINCPriority
    contact_type= $SNINCContactType
    notify= $SNINCNotify
    watch_list= $SNINCWatchlist
    service_offering= $SNINCServiceOffering
    u_production_impact= $SNINCProductionImpact
    category= $SNINCCategory
    subcategory= $SNINCSubcategory
    u_item= $SNINCItem
    assignment_group= $SNINCAssignmentGroup
    assigned_to= $SNINCAssignedTo
    short_description= $SNINCShortDescription
    description= $SNINCDescription
}
$SNCreateINCbodyjson = $SNCreateINCBody | ConvertTo-Json

#______________________________________________________________________________________
# POST to API
Try 
{
#______________________________________________________________________________________
# Send API request
$SNCreateIncResponse = Invoke-RestMethod -Method $SNMethodPost -Uri $SNINCAddress -Body $SNCreateINCbodyjson -TimeoutSec 100 -Headers $SNheaders -ContentType "application/json"
}
Catch 
{
Write-Host $_.Exception.ToString()
$error[0] | Format-List -Force
}
#______________________________________________________________________________________
# Verifying Incident created and show ID
IF ($SNCreateIncResponse.result.number -ne $null)
{
return $SNCreateIncResponse
}
ELSE
{
"Incident Not Created"
}
}

#______________________________________________________________________________________
#Import Create-Change Function
Function Create-Change($SNCCParentINC)
{
#______________________________________________________________________________________
#Specify Change Request Body
$SNCreateCHGbody = @{ #Create Body of the Post Request
    requested_by = $SNCHGRequestedBy
    category = $SNCHGCategory
    service_offering = $SNCHGServiceOffering
    reason = $SNCHGReason
    u_client_impact = $SNCHGClientImpact
    start_date = $SNCHGStartDate
    end_date = $SNCHGEndDate
    watch_list = $SNCHGWatchList
    parent = $SNCCParentINC
    urgency = $SNCHGUrgency
    risk = $SNCHGRisk
    type = $SNCHGType
    state = $SNCHGState
    assignment_group = $SNCHGAssignmentGroup
    assigned_to = $SNCHGAssignedTo
    short_description = $SNINCShortDescription
    description = $SNINCDescription
    justification = $SNCHGJustification
    change_plan = $SNCHGChangePlan
    test_plan = $SNCHGTestPlan
    backout_plan = $SNCHGBackoutPlan
}
$SNCreateCHGbodyjson = $SNCreateCHGbody | ConvertTo-Json

#______________________________________________________________________________________
# POST to API
Try 
{

#______________________________________________________________________________________
# Send API request
$SNCreateChangeResponse = Invoke-RestMethod -Method $SNMethodPost -Uri $SNCHGAddress -Body $SNCreateCHGbodyjson -TimeoutSec 100 -Headers $SNheaders -ContentType "application/json"
}
Catch 
{
Write-Host $_.Exception.ToString()
$error[0] | Format-List -Force
}

#______________________________________________________________________________________
# Pulling ticket ID from response
$SNCreateChangeID = $SNCreateChangeResponse.result.number
$SNCreateChangeSysID = $SNCreateChangeResponse.result.sys_id

#______________________________________________________________________________________
# Verifying Change created and show ID
IF ($SNCreateChangeID -ne $null)
{
return $SNCreateChangeResponse
}
ELSE
{
"Change Not Created"
}
}

#______________________________________________________________________________________
#Import Update-Change Function
Function Update-Change($SNCHGUpdateComments)
{
#______________________________________________________________________________________
# Specify request body
$SNUpdateCHGbody = @{ #Create Body of the Post Request
    comments= $SNCHGUpdateComments
}
$SNUpdateCHGbodyjson = $SNUpdateCHGbody | ConvertTo-Json

#______________________________________________________________________________________
# Send API request
$SNUpdateChangeResponse = Invoke-RestMethod -Method $SNMethodPatch -Uri "$SNCHGAddress\$SNChangeSysID" -Body $SNUpdateCHGbodyjson -TimeoutSec 100 -Headers $SNheaders -ContentType "application/json"
}

#______________________________________________________________________________________
#Import Update-Incident Function
Function Update-Incident($SNINCUpdateWorkNotesUpdate)
{

#______________________________________________________________________________________
# Specify request body
$SNUpdateINCbody = @{ #Create Body of the Post Request
    work_notes= $SNINCUpdateWorkNotesUpdate
}
$SNUpdateINCbodyjson = $SNUpdateINCbody | ConvertTo-Json

#______________________________________________________________________________________
# Send API request
$SNUpdateIncidentResponse = Invoke-RestMethod -Method $SNMethodPatch -Uri "$SNINCAddress\$SNIncidentSysID" -Body $SNUpdateINCbodyjson -TimeoutSec 100 -Headers $SNheaders -ContentType "application/json"
}

#______________________________________________________________________________________
#Import Update-ChangeClosePrep Function
Function Update-ChangeClosePrep($SNCHGUpdateComments, $SNCHGUpdateChangeSummary)
{
#______________________________________________________________________________________
# Specify request body
$SNUpdateCHGbody = @{ #Create Body of the Post Request
    comments= $SNCHGUpdateComments
    u_change_summary= $SNCHGUpdateChangeSummary
}
$SNUpdateCHGbodyjson = $SNUpdateCHGbody | ConvertTo-Json

#______________________________________________________________________________________
# Send API request
$SNUpdateChangeResponse = Invoke-RestMethod -Method $SNMethodPatch -Uri "$SNCHGAddress\$SNChangeSysID" -Body $SNUpdateCHGbodyjson -TimeoutSec 100 -Headers $SNheaders -ContentType "application/json"
}

#______________________________________________________________________________________
#Import Update-IncidentClosePrep Function
Function Update-IncidentClosePrep($SNINCUpdateWorkNotesUpdate, $SNINCUpdateCloseNotesUpdate)
{

#______________________________________________________________________________________
# Specify request body
$SNUpdateINCbody = @{ #Create Body of the Post Request
    work_notes= $SNINCUpdateWorkNotesUpdate
    close_notes= $SNINCUpdateCloseNotesUpdate
}
$SNUpdateINCbodyjson = $SNUpdateINCbody | ConvertTo-Json

#______________________________________________________________________________________
# Send API request
$SNUpdateIncidentResponse = Invoke-RestMethod -Method $SNMethodPatch -Uri "$SNINCAddress\$SNIncidentSysID" -Body $SNUpdateINCbodyjson -TimeoutSec 100 -Headers $SNheaders -ContentType "application/json"
}

#______________________________________________________________________________________
#Import Get-Incident Function
Function Get-Incident($SNGetIncidentSysID)
{
#______________________________________________________________________________________
# Build URI
$SNGetINCAddress = "$SNINCAddress/$SNGetIncidentSysID" + "?sysparm_fields=parent%2Ccaused_by%2Cwatch_list%2Cu_aging_category%2Cu_call_back_number%2Cupon_reject%2Csys_updated_on%2Cu_resolved_by_tier_1%2Cu_ud_parent%2Cu_resolved_within_1_hour%2Cu_routing_rule%2Capproval_history%2Cskills%2Cu_actual_resolution_date%2Cnumber%2Cu_related_incidents%2Cu_closure_category%2Cstate%2Csys_created_by%2Cknowledge%2Corder%2Cdelivery_plan%2Ccmdb_ci%2Cimpact%2Cu_requested_for%2Cactive%2Cpriority%2Cgroup_list%2Cbusiness_duration%2Cu_template%2Capproval_set%2Cwf_activity%2Cu_requested_by_phone%2Cshort_description%2Cu_itil_watch_list%2Cdelivery_task%2Ccorrelation_display%2Cwork_start%2Cu_ca_reference%2Cadditional_assignee_list%2Cnotify%2Cservice_offering%2Csys_class_name%2Cfollow_up%2Cclosed_by%2Creopened_by%2Cu_csv_comments%2Cu_planned_response_date%2Creassignment_count%2Cassigned_to%2Csla_due%2Cu_actual_response_date%2Cu_sla_met%2Cu_closure_ci%2Cu_reopen_count%2Cescalation%2Cupon_approval%2Cu_service_category%2Ccorrelation_id%2Cu_resolution_duration%2Cu_requested_by_name%2Cmade_sla%2Cu_requested_by_email%2Cu_item%2Cu_svc_desk_created%2Cresolved_by%2Cu_business_service%2Csys_updated_by%2Cuser_input%2Copened_by%2Csys_created_on%2Csys_domain%2Cu_quality_impact%2Cu_req_count%2Ccalendar_stc%2Cclosed_at%2Cu_relationship%2Cu_parent_incident%2Cu_comments_and_work_notes%2Cu_requested_by_not_found%2Cu_requested_by%2Cbusiness_service%2Cu_agile_incident_ref%2Cu_symptom%2Crfc%2Ctime_worked%2Cexpected_start%2Copened_at%2Cwork_end%2Creopened_time%2Cresolved_at%2Ccaller_id%2Cu_client%2Cwork_notes%2Csubcategory%2Cu_ah_incident%2Cclose_code%2Cassignment_group%2Cbusiness_stc%2Cdescription%2Cu_planned_resolved_date%2Ccalendar_duration%2Cu_on_hold_type%2Cu_source%2Cclose_notes%2Cu_closure_subcategory%2Cu_previous_assignment%2Csys_id%2Ccontact_type%2Curgency%2Cproblem_id%2Cu_itil_group_list%2Cu_response_duration%2Cu_best_number%2Ccompany%2Cactivity_due%2Cseverity%2Cu_production_impact%2Ccomments%2Capproval%2Cdue_date%2Csys_mod_count%2Csys_tags%2Clocation%2Ccategory"

#______________________________________________________________________________________
# Specify request body
$SNGetINCbody = @{ }
$SNGetINCbodyjson = $SNGetINCbody | ConvertTo-Json

#______________________________________________________________________________________
# Send API request
$SNGetIncidentResponse = Invoke-RestMethod -Method $SNMethodGet -Headers $SNHeaders -Uri $SNGetINCAddress

Return $SNGetIncidentResponse
}

#______________________________________________________________________________________
#Import Get-Change Function
Function Get-Change($SNGetChangeSysID)
{

#______________________________________________________________________________________
# Build URI
$SNGetCHGAddress = "$SNCHGAddress/$SNGetChangeSysID" + "?sysparm_fields=reason%2Cparent%2Cwatch_list%2Cu_aging_category%2Cproposed_change%2Cu_notification_form%2Cu_ah_change%2Cu_call_back_number%2Cupon_reject%2Csys_updated_on%2Ctype%2Cu_resolved_by_tier_1%2Cu_ud_parent%2Cu_resolved_within_1_hour%2Cu_routing_rule%2Capproval_history%2Cskills%2Ctest_plan%2Cu_actual_resolution_date%2Cnumber%2Cu_related_incidents%2Ccab_delegate%2Crequested_by_date%2Cu_business_impact%2Cu_validation_impact%2Cci_class%2Cstate%2Csys_created_by%2Cknowledge%2Corder%2Cphase%2Cdelivery_plan%2Ccmdb_ci%2Cimpact%2Cu_requested_for%2Cactive%2Cu_change_summary%2Cpriority%2Ccab_recommendation%2Cproduction_system%2Creview_date%2Cu_record_producer%2Crequested_by%2Cgroup_list%2Cbusiness_duration%2Cu_template%2Cchange_plan%2Capproval_set%2Cwf_activity%2Cimplementation_plan%2Cu_requested_by_phone%2Cstatus%2Cend_date%2Cshort_description%2Cu_itil_watch_list%2Cdelivery_task%2Ccorrelation_display%2Cwork_start%2Cu_ca_reference%2Coutside_maintenance_schedule%2Cadditional_assignee_list%2Cservice_offering%2Csys_class_name%2Cfollow_up%2Cclosed_by%2Cu_technical_impact%2Cu_depl_pkg_requested%2Cu_planned_response_date%2Creview_status%2Creassignment_count%2Cstart_date%2Cassigned_to%2Csla_due%2Cu_actual_response_date%2Cu_sla_met%2Cu_reopen_count%2Cescalation%2Cupon_approval%2Cu_service_category%2Ccorrelation_id%2Cu_resolution_duration%2Cu_requested_by_name%2Cmade_sla%2Cbackout_plan%2Cu_requested_by_email%2Cconflict_status%2Cu_item%2Cu_business_service%2Csys_updated_by%2Cuser_input%2Copened_by%2Csys_created_on%2Cu_cab_approval%2Csys_domain%2Cu_quality_impact%2Cu_req_count%2Cclosed_at%2Cu_relationship%2Creview_comments%2Cu_comments_and_work_notes%2Cu_requested_by_not_found%2Cu_requested_by%2Cbusiness_service%2Cu_symptom%2Ctime_worked%2Cexpected_start%2Copened_at%2Cwork_end%2Cphase_state%2Ccab_date%2Cwork_notes%2Csubcategory%2Cassignment_group%2Cdescription%2Cu_planned_resolved_date%2Cu_client_impact%2Ccalendar_duration%2Cu_on_hold_type%2Cclose_notes%2Csys_id%2Ccontact_type%2Ccab_required%2Cu_cab_yes%2Curgency%2Cscope%2Cu_itil_group_list%2Cu_response_duration%2Ccompany%2Cjustification%2Cactivity_due%2Ccomments%2Capproval%2Cdue_date%2Csys_mod_count%2Csys_tags%2Cconflict_last_run%2Crisk%2Clocation%2Ccategory%2Ccaused_by%2Cu_closure_category%2Cnotify%2Creopened_by%2Cu_csv_comments%2Cu_closure_ci%2Cu_svc_desk_created%2Cresolved_by%2Ccalendar_stc%2Cu_parent_incident%2Cu_agile_incident_ref%2Crfc%2Creopened_time%2Cresolved_at%2Ccaller_id%2Cu_client%2Cu_ah_incident%2Cclose_code%2Cbusiness_stc%2Cu_source%2Cu_closure_subcategory%2Cu_previous_assignment%2Cproblem_id%2Cu_best_number%2Cseverity%2Cu_production_impact"

#______________________________________________________________________________________
# Specify request body
$SNGetCHGbody = @{ }
$SNGetCHGbodyjson = $SNGetCHGbody | ConvertTo-Json

#______________________________________________________________________________________
# Send API request
$SNGetChangeResponse = Invoke-RestMethod -Method $SNMethodGet -Headers $SNHeaders -Uri $SNGetCHGAddress

Return $SNGetChangeResponse
}

#______________________________________________________________________________________
#Import Update-ChangeClose Function
Function Update-ChangeClose()
{
#______________________________________________________________________________________
# Specify request body
$SNUpdateCHGbody = @{ #Create Body of the Post Request
    state= '3'
    u_reopen_count= '1'
}
$SNUpdateCHGbodyjson = $SNUpdateCHGbody | ConvertTo-Json

#______________________________________________________________________________________
# Send API request
$SNUpdateChangeResponse = Invoke-RestMethod -Method $SNMethodPatch -Uri "$SNCHGAddress\$SNChangeSysID" -Body $SNUpdateCHGbodyjson -TimeoutSec 100 -Headers $SNheaders -ContentType "application/json"
}

#______________________________________________________________________________________
#Import Update-IncidentClose Function
Function Update-IncidentClose()
{

#______________________________________________________________________________________
# Specify request body
$SNUpdateINCbody = @{ #Create Body of the Post Request
    state= '6'
}
$SNUpdateINCbodyjson = $SNUpdateINCbody | ConvertTo-Json

#______________________________________________________________________________________
# Send API request
$SNUpdateIncidentResponse = Invoke-RestMethod -Method $SNMethodPatch -Uri "$SNINCAddress\$SNIncidentSysID" -Body $SNUpdateINCbodyjson -TimeoutSec 100 -Headers $SNheaders -ContentType "application/json"
}

}
#______________________________________________________________________________________
#Create Install and Shutdown Scripts
#______________________________________________________________________________________

#______________________________________________________________________________________
#Define InstallUpdateScript File Name Var
$InstallUpdatesScript = "InstallUpdates.ps1"

#______________________________________________________________________________________
#Build Install Script File
@"
#______________________________________________________________________________________
# Do not Edit Below! Do not Edit Below! Do not Edit Below! Do not Edit Below!
#______________________________________________________________________________________


#______________________________________________________________________________________
#Varibles
`$MasterVMHostName = hostname
`$ScriptDate = Get-Date -Format MM-dd-yyyy
`$ScriptDate2 = Get-Date -UFormat %m/%d/%y

#______________________________________________________________________________________
#Define Functions
Function Get-Software  
{

    [OutputType('System.Software.Inventory')]

    [Cmdletbinding()] 

    Param( 
            [Parameter(ValueFromPipeline=`$True,ValueFromPipelineByPropertyName=`$True)] 
            [String[]]`$Computername=`$env:COMPUTERNAME
         )            

    Begin { }

    Process  
    {      
        #______________________________________________________________________________________
        # Build Paths
        `$Paths  = @("SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Uninstall","SOFTWARE\\Wow6432node\\Microsoft\\Windows\\CurrentVersion\\Uninstall")         
        ForEach(`$Path in `$Paths) 
            { 
                Write-Verbose  "Checking Path: `$Path"
                #______________________________________________________________________________________
                # Create an instance of the Registry Object and open the HKLM base key 
                Try  {`$reg=[microsoft.win32.registrykey]::OpenRemoteBaseKey('LocalMachine',`$Computer,'Registry64')} 
                Catch  
                       { Write-Error `$_ 
                       Continue 
                       } 
                
                #______________________________________________________________________________________
                # Drill down into the Uninstall key using the OpenSubKey Method 
                Try  {
                        `$regkey=`$reg.OpenSubKey(`$Path)  

                        #______________________________________________________________________________________
                        # Retrieve an array of string that contain all the subkey names 
                        `$subkeys=`$regkey.GetSubKeyNames()
                        #______________________________________________________________________________________      
                        # Open each Subkey and use GetValue Method to return the required  values for each 
                        ForEach (`$key in `$subkeys)
                            {   
                                Write-Verbose "Key: `$Key"
                                `$thisKey=`$Path+"\\"+`$key 
                                Try {  
                                        `$thisSubKey=`$reg.OpenSubKey(`$thisKey) 
                                        #______________________________________________________________________________________  
                                        # Prevent Objects with empty DisplayName 
                                        `$DisplayName =  `$thisSubKey.getValue("DisplayName")
                                        If (`$DisplayName  -AND `$DisplayName  -notmatch '^Update  for|rollup|^Security Update|^Service Pack|^HotFix') 
                                        {
                                            `$Date = `$thisSubKey.GetValue('InstallDate')
                                            If (`$Date) 
                                                {
                                                    Try {`$Date = [datetime]::ParseExact(`$Date, 'yyyyMMdd', `$Null)} 
                                                    Catch{Write-Warning "`$(`$Computer): `$_ <`$(`$Date)>"
                                                    `$Date = `$Null}
                                                } 
                                            #______________________________________________________________________________________
                                            # Create New Object with empty Properties 
                                            `$Publisher =  Try {`$thisSubKey.GetValue('Publisher').Trim()} 
                                            Catch {`$thisSubKey.GetValue('Publisher')}

                                            `$Version = Try {`$thisSubKey.GetValue('DisplayVersion').TrimEnd(([char[]](32,0)))} 
                                            Catch {`$thisSubKey.GetValue('DisplayVersion')}

                                            `$UninstallString =  Try {`$thisSubKey.GetValue('UninstallString').Trim()} 
                                            Catch {`$thisSubKey.GetValue('UninstallString')}

                                            `$InstallLocation =  Try {`$thisSubKey.GetValue('InstallLocation').Trim()} 
                                            Catch {`$thisSubKey.GetValue('InstallLocation')}

                                            `$InstallSource =  Try {`$thisSubKey.GetValue('InstallSource').Trim()} 
                                            Catch {`$thisSubKey.GetValue('InstallSource')}

                                            `$HelpLink = Try {`$thisSubKey.GetValue('HelpLink').Trim()} 
                                            Catch {`$thisSubKey.GetValue('HelpLink')}

                                            #______________________________________________________________________________________
                                            # Fill Objects with Data
                                            `$Object = [pscustomobject]@{
                                            Computername = `$Computer
                                            DisplayName = `$DisplayName
                                            Version  = `$Version
                                            InstallDate = `$Date
                                            Publisher = `$Publisher
                                            UninstallString = `$UninstallString
                                            InstallLocation = `$InstallLocation
                                            InstallSource  = `$InstallSource
                                            HelpLink = `$thisSubKey.GetValue('HelpLink')
                                            EstimatedSizeMB = [decimal]([math]::Round((`$thisSubKey.GetValue('EstimatedSize')*1024)/1MB,2))}

                                            `$Object.pstypenames.insert(0,'System.Software.Inventory')
                                            Write-Output `$Object
                                        }
                                    } Catch {Write-Warning "`$Key : `$_"}   
                            }
                     } Catch  {}   
                    `$reg.Close() 
            } 
    } 
}  

#______________________________________________________________________________________
# Start Services Function
function Start-VDIservices (`$SVCname) {
    Set-Service -Name `$svcname -StartupType Automatic
    write-host "Changed Service startyp type for `$Svcname to the following"
    Get-Service -Name `$svcname | Select-Object StartType
    Start-Service -Name `$svcname
    Write-Host "`$SVCname running status is:"
    Get-Service -Name `$svcname | Select-Object Status
    Do {
    `$svc = Get-Service -Name `$svcname
    Start-Sleep 2
    Write-Host "`$SVCname running status is:"
    `$svc
    } While ( `$svc.Status -ne "Running" )
    }

#______________________________________________________________________________________
# Capture installed Software before updating
`$InstalledSoftwareBefore = get-software | Format-Table | Out-File -Filepath "$LogLocation\`$ScriptDate\InstalledSoftwareBefore.`$MasterVMHostName.txt"
write-host "Capture Installed Software Prior to update and save to Logs Folder"

#______________________________________________________________________________________
#Check to see if machine is managed by SCCM
`$CheckSCCM = get-wmiobject win32_Service | Where-Object {`$_.Name -eq "CCMexec"} | Select-Object Name
write-host "Check if VM is managed by SCCM"

#______________________________________________________________________________________
#If Else statement based on if VM is managed by SCCM
if(`$CheckSCCM.Name -eq 'CCMexec')
{
    write-host "VM is managed by SCCM"
    #______________________________________________________________________________________
    #Check Status of SCCM service
    `$SCCMStatus = get-service -name CCMexec | Select-Object Name,Status,Starttype
    if(`$SCCMStatus.Status -eq 'Stopped')
    {   
        #______________________________________________________________________________________
        #Set SCCM to Automatic and Start Service
        Start-VDIservices CCMexec
        write-host "Start SCCM Services"
    }

    #______________________________________________________________________________________
    #Start SCCM Patching
    `$AppEvalState0 = "0" 
    `$AppEvalState1 = "1" 
    `$Application = (Get-WmiObject -Namespace "root\ccm\clientSDK" -Class CCM_SoftwareUpdate | Where-Object { `$_.EvaluationState -like "*`$(`$AppEvalState0)*" -or `$_.EvaluationState -like "*`$(`$AppEvalState1)*"})
    Invoke-WmiMethod -Class CCM_SoftwareUpdatesManager -Name InstallUpdates -ArgumentList (,`$Application) -Namespace root\ccm\clientsdk
    write-host "Invoke SCCM install of Updates"

    `$SCCMPatternDate = get-date -Format MM-dd-yyyy

    ([wmiclass]‘root\ccm:SMS_Client’).TriggerSchedule(‘{00000000-0000-0000-0000-000000000113}’)
    `$SCCMPatternData = Select-String -Path C:\Windows\CCM\Logs\UpdatesDeployment.log -Pattern `$SCCMPatternDate  | Select-Object -Last 12
    While((`$SCCMPatternData -notmatch "1881") -eq "True")
    {
    ([wmiclass]‘root\ccm:SMS_Client’).TriggerSchedule(‘{00000000-0000-0000-0000-000000000113}’)
    `$SCCMPatternData = Select-String -Path C:\Windows\CCM\Logs\UpdatesDeployment.log -Pattern `$SCCMPatternDate  | Select-Object -Last 12
    }
    
}
else
#______________________________________________________________________________________
#Else statemet for installing Windows updates through windows updates, and making external calls to other scripts to update 
{
    #______________________________________________________________________________________
    #Check Status of Windows Update service
    `$WindowsUpdatateStatus = get-service -name wuauserv | Select-Object Name,Status,Starttype
    if(`$WindowsUpdatateStatus.Status -eq 'Stopped')
    { 
        #______________________________________________________________________________________
        #Set Windows Update to Automatic and Start Service
        Start-VDIservices wuauserv
        write-host "Started Windows Update Service"
    }

    #______________________________________________________________________________________
    #Install 3rd Party updates (Java, Flash, Adobe Reader, Chrome, Firefox)

    #______________________________________________________________________________________
    #Check to see if Adobe Flash is installed
    if ((Get-Software | Select-Object DisplayName) -like "*Flash*")
    {
        #______________________________________________________________________________________
        #Install Flash Updates
        Invoke-Expression -command "$ScriptLocation\$FlashUpdate"
        write-host "Installed Flash Update"
    }
    
    #______________________________________________________________________________________
    #Check to see if Adobe Acrobat is installed
    if ((Get-Software | Select-Object DisplayName) -like "*Acrobat Reader*")
    {
        #______________________________________________________________________________________
        #Install Adobe Updates
        Invoke-Expression -command "$ScriptLocation\$AdobeUpdate"
        write-host "Installed Adobe Acrobat Update"
    }
    
    #______________________________________________________________________________________
    #Check to see if Firefox is installed
    if ((Get-Software | Select-Object DisplayName) -like "*Mozilla Firefox*")
    {
        #______________________________________________________________________________________
        #Install Firefox Updates
        Invoke-Expression -command "$ScriptLocation\$FirefoxUpdate"
        write-host "Installed Firefox Update"
    }
    
    #______________________________________________________________________________________
    #Check to see if Java is installed
    if ((Get-Software | Select-Object DisplayName) -like "*Java*")
    {
        #______________________________________________________________________________________
        #Install Java Updates
        Invoke-Expression -command "$ScriptLocation\$JavaUpdate"
        write-host "Installed Java Update"
    }

    #______________________________________________________________________________________
    #Check to see if Chrome is installed
    if ((Get-Software | Select-Object DisplayName) -like "*Chrome*")
    {
        #______________________________________________________________________________________
        #Install Chrome Updates
        C:\"Program Files (x86)\Google\Update\GoogleUpdate.exe" /ua /installsource scheduler
        write-host "Installed Google Chrome Update"
    }

    #______________________________________________________________________________________
    #Prep for windows updates and Install updates
    #______________________________________________________________________________________
    #Hide Windows updates labled as "Preview"
    Hide-WindowsUpdate -Title "Preview*" -Confirm:`$false
    write-host "Hide Windows Updates labled as Preview"

    #______________________________________________________________________________________
    #Auto instal windows updates
    Get-WUInstall –AcceptAll -install
    write-host "Install Windows Updates"
    
    #______________________________________________________________________________________
    #Start Sleep timer for update installs.
    Start-Sleep -Seconds `$SleepTimeInS
    write-host "Wait for Windows Updates to install"

}

    #______________________________________________________________________________________
    # Check Version of OS
    `$winVersion = "0"
    `$OSVersion = [System.Environment]::OSVersion.Version
    if (`$OSversion.Major -eq "10" -and `$OSversion.Minor -eq "0"){ `$winVersion = "10" }
    elseif (`$OSversion.Major -eq "6" -and `$OSversion.Minor -eq "2") { `$winVersion = "8" }
    elseif (`$OSversion.Major -eq "6" -and `$OSversion.Minor -eq "1") { `$winVersion = "7" }
    Write-host "Windows Version is `$winVersion"
    `$Major = `$OSversion.Major
    `$Minor = `$OSversion.Minor
    `$Build = `$OSversion.Build
    `$Revision = `$OSversion.Revision

    `$CorporateBuildVersion = Get-ItemProperty -Path $CorporateBuildRegistryKeyPath -Name $CorporateBuildRegistryKeyName | Select-Object -ExpandProperty $CorporateBuildRegistryKeyName
    `$WindowsVersion = "Windows `$winVersion"
    `$WindowsBuildNumber = "`$Major.`$Minor.`$Build.`$Revision" 
    `$WindowsRevision = Get-ItemProperty -Path 'HKLM:SOFTWARE\Microsoft\Windows NT\CurrentVersion\' -Name ReleaseID | Select-Object -ExpandProperty ReleaseID
    `$LastUpdateDate = `$ScriptDate2    
#______________________________________________________________________________________
# Capture installed Software before updating
`$InstalledSoftwareAfter = get-software | Format-Table | Out-File "$LogLocation\`$ScriptDate\InstalledSoftwareAfter.`$MasterVMHostName.txt"
write-host "Capture Installed Software After update and save to Logs Folder"

#______________________________________________________________________________________
# Check if Custom REG key is created if not create Last Update Date REG key with 1/1/2000 date
if ( -not (Test-Path -Path $CustomRegPath)) {New-Item -path $CustomRegPath -Force | New-ItemProperty -Name  "$RegNameLastUpdateDate" -Value "01/01/2000" -Force | Out-Null}
write-host "Check if Custom key has been created and if not create registry key and set value to 1-1-2000"

#______________________________________________________________________________________
# Get Last Update Data value
`$LastUpdateDateValue = Get-ItemProperty -Path $CustomRegPath -Name "$RegNameLastUpdateDate" | Select-Object -ExpandProperty "$RegNameLastUpdateDate"
write-host "Get Last Update date from custom Registry key, Last Update date was `$LastUpdateDateValue"

#______________________________________________________________________________________
# Get Windows Udates installed since last update. 
`$WindowsUpdates = Get-HotFix | Where-Object -Property InstalledOn -GT `$LastUpdateDateValue | Out-File -FilePath "$LogLocation\`$ScriptDate\WindowsUpdates.`$MasterVMHostName.txt"
write-host "Create TXT file with Installed updates as of `$LastUpdateDateValue"

#______________________________________________________________________________________
#Set Custom Registry Values
Get-Item -path $CustomRegPath | New-ItemProperty -Name "$RegNameCorporateBuildVersion" -Value "`$CorporateBuildVersion" -Force | Out-Null
Get-Item -path $CustomRegPath | New-ItemProperty -Name "$RegNameWindowsVersion" -Value "`$WindowsVersion" -Force | Out-Null
Get-Item -path $CustomRegPath | New-ItemProperty -Name "$RegNameWindowsBuildNumber" -Value "`$WindowsBuildNumber" -Force | Out-Null
Get-Item -path $CustomRegPath | New-ItemProperty -Name "$RegNameWindowsRevision" -Value "`$WindowsRevision" -Force | Out-Null
Get-Item -Path $CustomRegPath | New-ItemProperty -Name "$RegNameLastUpdateDate" -Value "`$ScriptDate2" -Force | Out-Null
write-host "Custom Registry values on Master Image"

"@  |  Set-Content "$ScriptLocation\$InstallUpdatesScript"
Write-Host "Created Install Script and saved it to location $ScriptLocation\$InstallUpdatesScript"

#______________________________________________________________________________________
#Define ShutdownScript File Name Var
$ShutdownScript = "ShutDownScriptTest.ps1"

#______________________________________________________________________________________
#Build Shutdown Script File
@"
    #______________________________________________________________________________________
    # Check Version of OS
    `$winVersion = "0"
    `$OSVersion = [System.Environment]::OSVersion.Version
    if (`$OSversion.Major -eq "10" -and `$OSversion.Minor -eq "0"){ `$winVersion = "10" }
    elseif (`$OSversion.Major -eq "6" -and `$OSversion.Minor -eq "2") { `$winVersion = "8" }
    elseif (`$OSversion.Major -eq "6" -and `$OSversion.Minor -eq "1") { `$winVersion = "7" }
    Write-host "Windows Version is `$winVersion"
    
    #______________________________________________________________________________________
    # Varibles
    `$ScriptDate = Get-Date -Format MM-dd-yyyy
    `$MasterVMHostName = hostname
    `$RunOptimizer = $RunOptimizer
    `$RunSEP = $RunSEP
    `$defragbat = "$ScriptLocation\dfrag.bat"
    `$OptimizerLocation = "$CloneToolsLocation\$VMwareOptimizerName"
    `$OptimizerTemplate = "$CloneToolsLocation\$OptimizerTemplateNamingPrefix`${winVersion}.xml"
    `$OptimizerErrorLogLocation = "$LogLocation\$ScriptDate\OptimizationErrorLog.`$MasterVMHostName.txt"
    `$OptimizerbatFile = "$ScriptLocation\Optimizer.`$MasterVMHostName.bat"

    #______________________________________________________________________________________
    # Define functions

    #______________________________________________________________________________________
    # Start Services Function
    function Start-VDIservices (`$SVCname) {
    Set-Service -Name `$svcname -StartupType Automatic
    Start-Service -Name `$svcname
    Do {
    `$svc = Get-Service -Name `$svcname
    Start-Sleep 2
    } While ( `$svc.Status -ne "Running" )
    }

    #______________________________________________________________________________________
    # Stop Services Function
    function Stop-VDIservices (`$SVCname) {
    Stop-Service -Name `$svcname
    Do {
    `$svc = Get-Service -Name `$svcname
    Start-Sleep 2
    } While ( `$svc.Status -ne "Stopped" )
    Set-Service -Name `$svcname -StartupType disabled
    }

    #______________________________________________________________________________________
    # Do while proc Like Function
    function DoWhile-LikeProcName (`$Process) {
    Do
    {
    "Processes `$Process is still running"
    `$proc = Get-Process
    start-sleep 10
    } While (`$proc.name -like "*`$Process*")
    }

    #______________________________________________________________________________________
    # Do while Proc Equals Function
    function DoWhile-EQProcName (`$Process) {
    Do
    {
    "Processes `$Process is still running"
    `$proc = Get-Process
    start-sleep 10
    } While (`$proc.name -eq "`$Process")
    }    
    
    #______________________________________________________________________________________
    # Run Disk Cleanup to remove temp files, empty recycle bin and remove other unneeded files
    Start-Process -Filepath "c:\windows\system32\cleanmgr" -argumentlist '/sagerun:1'
    Write-host "Running Disk Cleanup"

    #______________________________________________________________________________________
    # Check status of Disk Cleanup
    DoWhile-LikeProcName cleanmgr

    #______________________________________________________________________________________
    # Check status of Dfrag Service if stopped start.
    `$DfragStatus = get-service -name defragsvc | Select-Object Name,Status,Starttype
    if(`$DfragStatus.Status -eq 'Stopped')
    {
    #______________________________________________________________________________________
    # Start Defrag Service
    Start-VDIservices defragsvc
    Write-host "Start Defrag Service"
    }

    #______________________________________________________________________________________
    # Create Temp Defrag BAT
    "defrag c: /U /V" | Set-Content `$defragbat
    Write-host "Create Defrag Temp .bat file"

    #______________________________________________________________________________________
    # Start Defrag BAT
    Start-Process `$defragbat
    Write-host "Running Defrag Cleanup"

    #______________________________________________________________________________________
    # Check status of Disk Defrag
    DoWhile-EQProcName Defrag

    #______________________________________________________________________________________
    # Remove Temp Defrag BAT
    Remove-Item -Path `$defragbat
    Write-host "Delete Defrag Temp .bat file"

    #______________________________________________________________________________________
    # Stop Defrag Service
    Stop-VDIservices defragsvc
    Write-host "Stop Defrag Service"

    #______________________________________________________________________________________
    # Pre-complile .NET framework Assemblies
    Start-Process -Filepath "C:\Windows\Microsoft.NET\Framework\v4.0.30319\ngen.exe" -argumentlist 'update','/force'
    New-ItemProperty -Name verbosestatus -Path "HKLM:\Software\Microsoft\Windows\CurrentVersion\Policies\System" -Force -PropertyType DWORD -Value "1"
    Write-host "Running a Pre-Compile of .NET Framework"

    #______________________________________________________________________________________
    # Check status of the .NET recompile
    DoWhile-LikeProcName ngen


    #______________________________________________________________________________________ 
    # Run SEP update and Cleanup
    if(`$RunSEP -eq '1')
    {
    Write-host "Starting the SEP Update and Cleanup Process"
    #______________________________________________________________________________________
    # Check to see if SEP is installed and prep for Recompose
    #______________________________________________________________________________________
    # If Check for SEP install
    `$CheckSEP = get-service -name SepMasterService | Select-Object Name,Status,Starttype
    if(`$CheckSEP.Name -eq 'SepMasterService')
    {

    #______________________________________________________________________________________
    # Force SEP checking and Update
    Start-Process -Filepath "C:\Program Files (x86)\Symantec\Symantec Endpoint Protection\smc.exe" -argumentlist '-updateconfig'
    Write-host "Update SEP Config"

    #______________________________________________________________________________________ 
    # Check status of SEP Update
    DoWhile-LikeProcName doscan

    #______________________________________________________________________________________ 
    # Run full SEP System scan
    Start-Process -Filepath "C:\Program Files (x86)\Symantec\Symantec Endpoint Protection\doScan.exe" -argumentlist '/scanname "Full System Scan"'
    Write-host "Run Full System SEP Scan"

    #______________________________________________________________________________________ 
    # Check status of SEP Scan
    DoWhile-LikeProcName doscan

    #______________________________________________________________________________________ 
    # Force SEP checking and Update
    Start-Process -Filepath "C:\Program Files (x86)\Symantec\Symantec Endpoint Protection\smc.exe" -argumentlist '-updateconfig'
    Write-host "Update SEP and force Checkin to report the recent completed Scan"

    #______________________________________________________________________________________ 
    # Check status of SEP Update
    DoWhile-LikeProcName doscan

    #______________________________________________________________________________________
    # Check status of Symantec Endpoint Protection Service.
    `$SEPStatus = get-service -name SepMasterService | Select-Object Name,Status,Starttype
    if(`$SEPStatus.Status -eq 'Running')
    {
        #______________________________________________________________________________________
        # Stop SEP Service and Prep for Clone
        Start-Process -Filepath "C:\Program Files (x86)\Symantec\Symantec Endpoint Protection\smc.exe" -argumentlist '-stop'
        Write-host "Stopped SEP Service"
    }

        `$filesToClean = @("C:\sephwid.xml","C:\communicator.dat","C:\Program Files\Common Files\Symantec Shared\HWID\sephwid.xml","C:\Program Files\Common Files\Symantec Shared\HWID\communicator.dat","C:\Documents and Settings\All Users\Application Data\Symantec\Symantec Endpoint Protection\CurrentVersion\Data\Config\sephwid.xml","C:\Documents and Settings\All Users\Application Data\Symantec\Symantec Endpoint Protection\CurrentVersion\Data\Config\communicator.dat","C:\Documents and Settings\All Users\Application Data\Symantec\Symantec Endpoint Protection\PersistedData\sephwid.xml","C:\Documents and Settings\All Users\Application Data\Symantec\Symantec Endpoint Protection\PersistedData\communicator.dat","C:\ProgramData\Symantec\Symantec Endpoint Protection\PersistedData\sephwid.xml","C:\ProgramData\Symantec\Symantec Endpoint Protection\PersistedData\communicator.dat","C:\Users\All Users\Symantec\Symantec Endpoint Protection\PersistedData\sephwid.xml","C:\Users\All Users\Symantec\Symantec Endpoint Protection\PersistedData\communicator.dat","C:\Windows\Temp\sephwid.xml","C:\Windows\Temp\communicator.dat","C:\Documents and Settings\*\Local Settings\Temp\sephwid.xml","CC:\Documents and Settings\*\Local Settings\Temp\communicator.dat","C:\Users\*\AppData\Local\Temp\sephwid.xml","C:\Users\*\AppData\Local\Temp\communicator.dat")

        foreach(`$file in `$filesToClean)
        {
            if(Test-Path `$file) {Remove-Item -Path `$file -force}
        }
        Write-host "Removed the following files: `$FilesToClean"

        #______________________________________________________________________________________ 
        # Remove SEP Reg Keys

        #______________________________________________________________________________________ 
        # Remove 32bit SEP Reg Keys
        if (Test-Path -Path "HKLM:\SOFTWARE\Symantec\Symantec Endpoint Protection\SMC\SYLINK\SyLink")
        {
            Remove-ItemProperty -Path "HKLM:\SOFTWARE\Symantec\Symantec Endpoint Protection\SMC\SYLINK\SyLink" -Name "ForceHardwareKey"
            Remove-ItemProperty -Path "HKLM:\SOFTWARE\Symantec\Symantec Endpoint Protection\SMC\SYLINK\SyLink" -Name "HardwareID"
            Remove-ItemProperty -Path "HKLM:\SOFTWARE\Symantec\Symantec Endpoint Protection\SMC\SYLINK\SyLink" -Name "HostGUID"
            Write-host "Removed SEP 32 Bit Reg Keys"
        }
        if (Test-Path -Path "HKLM:\SOFTWARE\WOW6432Node\Symantec\Symantec Endpoint Protection\SMC\SYLINK\SyLink")
        {
            Remove-ItemProperty -Path "HKLM:\SOFTWARE\WOW6432Node\Symantec\Symantec Endpoint Protection\SMC\SYLINK\SyLink" -Name "ForceHardwareKey"
            Remove-ItemProperty -Path "HKLM:\SOFTWARE\WOW6432Node\Symantec\Symantec Endpoint Protection\SMC\SYLINK\SyLink" -Name "HardwareID"
            Remove-ItemProperty -Path "HKLM:\SOFTWARE\WOW6432Node\Symantec\Symantec Endpoint Protection\SMC\SYLINK\SyLink" -Name "HostGUID"
            Write-host "Removed SEP 64 Bit Reg Keys"
            }
    
    }
    Write-host "Finished SEP Cleanup and Prep for Cloneing"
    }

    #______________________________________________________________________________________ 
    # Run VMware optimizer
    if(`$RunOptimizer -eq '1')
    {
    Write-host "Prepairing to run VMware Optimizer"

#______________________________________________________________________________________ 
# Create Temp Bat file to optimize OS
`@"
            start /wait `$OptimizerLocation -t `$OptimizerTemplate -r $LogLocation\`$ScriptDate\
            start /wait `$OptimizerLocation -o recommended -t `$OptimizerTemplate -v > `$OptimizerErrorLogLocation 2>&1
            start /wait `$OptimizerLocation -t `$OptimizerTemplate -r $LogLocation\`$ScriptDate\
"`@ | Set-Content `$OptimizerbatFile

    Write-host "Built Temporary VMware Optimizer .BAT file based on Template `$OptimizerTemplate"
    #______________________________________________________________________________________ 
    # Run BAT file to optimize OS
    Start-Process `$OptimizerbatFile
    Write-host "Running VMware Optimizer .BAT based on Template `$OptimizerTemplate"

    #______________________________________________________________________________________ 
    # Check status of the Optimize and pause
    DoWhile-EQProcName CMD

    #______________________________________________________________________________________ 
    # Remove Temp BAT file
    Remove-Item -Path `$OptimizerbatFile
    Write-host "Removing Temporary VMware Optimizer .BAT file"

    }
    #______________________________________________________________________________________
    # Check if VM is managed by SCCM
    `$CheckSCCM = get-wmiobject win32_Service | Where-Object {`$_.Name -eq "CCMexec"} | Select-Object Name
    Write-host "Check if SCCM service is installed"

    #______________________________________________________________________________________
    # If Else statement based on if VM is managed by SCCM
    if(`$CheckSCCM.Name -eq 'CCMexec')
    {
    Write-host "VM is managed by SCCM"

    #______________________________________________________________________________________
    # Check status of SCCM Service.
    `$SCCMStatus = get-service -name "CCMexec" | Select-Object Status
    Write-host "Check if SCCM service CCMexec is running"
    if(`$SCCMStatus.Status -eq 'Running')
        {   
            Stop-VDIservices CCMexec
            Write-host "Stop SCCM Service CCMexec"
        }
    
    #______________________________________________________________________________________
    # Check status of SCCM Service if stopped start.
    `$AdaptivaStatus = get-service -name "adaptivaclient" | Select-Object Status
    Write-host "Check if Adaptiva Client is installed"
    if(`$AdaptivaStatus.Status -eq 'Running')
        {
            Stop-VDIservices adaptivaclient
            Write-host "Stop Adaptiva Services"
        }
    }

    #______________________________________________________________________________________
    # Check status of Windows Update Service.
    `$WindowsUpdatateStatus = get-service -name wuauserv | Select-Object Name,Status,Starttype
    if(`$WindowsUpdatateStatus.Status -eq 'Running')
    {
        #______________________________________________________________________________________
        # Stop Windows Update Service
        Stop-VDIservices wuauserv
        Write-host "Stop Windows Update Services"
    }

    #______________________________________________________________________________________
    # Clean out widows downloads cache
    if (Test-Path "C:\Windows\SoftwareDistribution\Download\")
    {
    Remove-Item -Path "C:\Windows\SoftwareDistribution\Download\*" -force -Recurse
    Write-host "Clean Out Windows Update Download Cache"
    }

    #______________________________________________________________________________________
    # Clean out Volume Shadow Copys
    #.\vssadmin.exe delete shadows /all
    #Write-host "Clean out Volume Shadow Copys"

    #______________________________________________________________________________________
    # Clean out Windows PreFetch
    if (Test-Path "c:\Windows\Prefetch\")
    {
    Remove-item -Path "c:\Windows\Prefetch\*.*" -Force
    Write-host "Remove windows Prefetch files"
    }

    #______________________________________________________________________________________
    # Check status if AppVolumes is installed and stop the service.
    `$AppVStatus = get-service | Select-Object Name,Status,Starttype
    if(`$AppVStatus.Name -eq 'svservice')
    {
    Stop-Service -Name svservice
    Write-host "Stop AppVolumes Service"
    }

    #______________________________________________________________________________________
    # Clear all event logs
    Get-EventLog -List
    `$Logs = Get-EventLog -List | ForEach {`$_.Log}
    `$Logs | ForEach {Clear-EventLog -Log `$_ }
    Write-host "Cleared Windows Event Logs"
    Get-EventLog -List


    #______________________________________________________________________________________
    # Release IP, Flush DNS and Shutdown VM
    ipconfig /release
    ipconfig /flushdns
    Stop-Computer -Force


"@ |  Set-Content "$ScriptLocation\$ShutdownScript"
Write-Host "Created Shutdown Script and saved it to location $ScriptLocation\$ShutdownScript"

#______________________________________________________________________________________
#Define Varibles for HVPool Extra varibles will be used in later revisions
Clear-Variable HVPool* -Scope Global

$HVPoolNAME=@()
$HVPoolMSTR=@()
$HVPoolProvisionType=@()
$HVPoolType=@()
$HVPoolAssign=@()
$HVPoolDC=@()
$HVPoolvCenter=@()

#______________________________________________________________________________________
#Connect to each Connection Server and get pool and vCenter info
foreach($HVServer in $HVServers)
{
#______________________________________________________________________________________
#Connect to Connection Server
Connect-HVServer $HVServer -Credential $VMwareSVCCreds
Write-Host "Connected to $HVServer"

#______________________________________________________________________________________
#Get Pool info
$pools = Get-HVPool
foreach($pool in $pools) {
$pool = $pool | get-hvpoolspec | ConvertFrom-Json
$HVPoolMSTR += $Pool.AutomatedDesktopSpec.virtualCenterProvisioningSettings.VirtualCenterProvisioningData.parentVm
$HVPoolvCenter += $Pool.AutomatedDesktopSpec.virtualCenter

#______________________________________________________________________________________
#Disconnect from Connection Server
Disconnect-HVServer $hvserver -Confirm:$false
Write-Host "Disconnected from $hvServer"
}
$HVPoolUniquevCenter = $HVPoolvCenter | select -Unique
}

#______________________________________________________________________________________
#Connect to vCenters
foreach ($vCenterServer in $HVPoolUniquevCenter) 
{
    Connect-VIServer $vCenterServer -Credential $VMwareSVCCreds
    Write-Host "Connected to vCenter Server $vCenterServer"
}

#______________________________________________________________________________________
#Create Service Now Incident and Change Ticket
if($RunServiceNow -eq '1')
{

#______________________________________________________________________________________
#Create Service Now Incident
$SNCreateIncResponseReturn = Create-Incident

#______________________________________________________________________________________
#Create Service Now Incident Varibles
$SNIncidentID = $SNCreateIncResponseReturn.result.number
$SNIncidentSysID = $SNCreateIncResponseReturn.result.sys_id

Write-Host "Created Service Now Incident $SNIncidentID with SYSid $SNIncidentSysID !"

#______________________________________________________________________________________
#Create Service Now Change
$SNCreateChangeResponseReturn = Create-Change $SNIncidentSysID

#______________________________________________________________________________________
#Create Service Now Change Varibles
$SNChangeID = $SNCreateChangeResponseReturn.result.number
$SNChangeSysID = $SNCreateChangeResponseReturn.result.sys_id

Write-Host "Created Service Now Incident $SNChangeID with SYSid $SNChangeSYSID !"
}

Write-Host "About to Invoke Update"

#______________________________________________________________________________________
#Run Install and Shutdown Scripts on each of the VDI Master Images
#______________________________________________________________________________________

#______________________________________________________________________________________
#Connect to each HV Server and populate pool info
foreach($HVServer in $HVServers)
{
Connect-HVServer $HVServer -Credential $VMwareSVCCreds
Write-Host "Connected to $HVServer"

#______________________________________________________________________________________
#Get Pool info and connect to each VM
$pools = Get-HVPool
foreach($pool in $pools) 
{
Write-Host "Running Update Process on Pool $pool."

#______________________________________________________________________________________
#Get Pool info from JSON
$pool = $pool | get-hvpoolspec | ConvertFrom-Json
$HVPoolNAME = $Pool.Base.name
$HVPoolMSTR = $Pool.AutomatedDesktopSpec.virtualCenterProvisioningSettings.VirtualCenterProvisioningData.parentVm
$HVPoolProvisionType = $Pool.AutomatedDesktopSpec.provisioningType
$HVPoolType = $Pool.Type
$HVPoolAssign = $Pool.AutomatedDesktopSpec.userAssignment.UserAssignment
$HVPoolDC = $Pool.AutomatedDesktopSpec.virtualCenterProvisioningSettings.VirtualCenterProvisioningData.datacenter
$VMLine = $HVPoolMSTR
$VMLineDomain = "$VMline.$DomainName"
    
#______________________________________________________________________________________
#PowerOn each of the master VMs
get-vm $VMLine | Where-Object {$_.powerstate -eq "poweredoff"} | Start-VM
Start-Sleep 240

$MasterLogLocation = $LogLocation -Replace "C\:\\","\\$VMLine\C$\"
$MasterScriptLocation =  $ScriptLocation -Replace "C\:\\","\\$VMLine\C$\"
$MasterCloneToolsLocation =  $CloneToolsLocation -Replace "C\:\\","\\$VMLine\C$"
$MasterVDIToolsLocation = "\\$VMLine\C$\VDI_Tools"

#______________________________________________________________________________________
#Create VDI_Tools Folder Structure on VDI Master VM
Create-Path $MasterLogLocation\$ScriptDate
Create-Path $MasterScriptLocation

#______________________________________________________________________________________
#Copy Scripts from Share Location to VDI Master VM
Copy-Item -Path $ShareScriptLocation\$InstallUpdatesScript -Destination $MasterScriptLocation
write-host "Copied Installed Updates Script from Share to VDI Master Image $VMLine in location $MasterScriptLocation\$InstallUpdatesScript"
Copy-Item -Path $ShareScriptLocation\$ShutdownScript -Destination $MasterScriptLocation
write-host "Copied Shutdown Script from Share to VDI Master Image $VMLine in location $MasterScriptLocation\$ShutdownScript"

#______________________________________________________________________________________
#If Software Update Scripts exist on the Share copy to Master VM
If-TestPath $AdobeUpdate
If-TestPath $FlashUpdate
If-TestPath $JavaUpdate
If-TestPath $FirefoxUpdate

#______________________________________________________________________________________
#Copy CloneTools from Share to Master VM
Copy-Item -Path $ShareCloneToolsLocation -Destination $MasterVDIToolsLocation -Recurse -Force
Write-host "Copy Clone tools from Share Location to VDI Master VM $VMLine to location $MasterVDIToolsLocation"

#______________________________________________________________________________________
#Start the Install Updates Script on Master VM
Write-Host "About to kick off the Install Updates Script on VDI Master $VMLine"
Invoke-Command -ComputerName $VMLineDomain -filepath "$ScriptLocation\$InstallUpdatesScript"
Write-Host "Finished running the Install Updates Script on VDI Master $VMLine"

#______________________________________________________________________________________
#Restart the Master VM
Write-Host "Restarting VDI Master VM $VMLine"
Restart-Computer -ComputerName $VMLineDomain -Force -Wait

#______________________________________________________________________________________
#ReStart the Debug Logging
$ErrorActionPreference="SilentlyContinue"
Stop-Transcript | out-null
$ErrorActionPreference = "Continue"
Start-Transcript -Force -Path $LogLocation\$ScriptDate\WeeklyVDIUpdates.txt -Append
Write-Host "Start Debug Logs"

#______________________________________________________________________________________
# Get recompose delay based on if pool is Test or Prod
if($HVPoolNAME -like $TestPoolNC)
{
$RecomposeDate = (Get-Date).AddHours($RCTimeDelay1)
}else {$RecomposeDate = (Get-Date).AddHours($RCTimeDelay2)}

#______________________________________________________________________________________
# Set Varibles With registry Keys and inputs from Pool Info Collection
$CorporateBuildVersion = Invoke-Command -ComputerName "$VMLine" -ScriptBlock {Get-ItemProperty -Path $($args[0]) -Name $($args[1]) | Select-Object -ExpandProperty $($args[1])} -ArgumentList $CustomRegPath, $RegNameCorporateBuildVersion
$WindowsVersion = Invoke-Command -ComputerName "$VMLine" -ScriptBlock {Get-ItemProperty -Path $($args[0]) -Name $($args[1]) | Select-Object -ExpandProperty $($args[1])} -ArgumentList $CustomRegPath, $RegNameWindowsVersion
$WindowsBuildNumber = Invoke-Command -ComputerName "$VMLine" -ScriptBlock {Get-ItemProperty -Path $($args[0]) -Name $($args[1]) | Select-Object -ExpandProperty $($args[1])} -ArgumentList $CustomRegPath, $RegNameWindowsBuildNumber
$WindowsRevision = Invoke-Command -ComputerName "$VMLine" -ScriptBlock {Get-ItemProperty -Path $($args[0]) -Name $($args[1]) | Select-Object -ExpandProperty $($args[1])} -ArgumentList $CustomRegPath, $RegNameWindowsRevision
$LastUpdateDate = Invoke-Command -ComputerName "$VMLine" -ScriptBlock {Get-ItemProperty -Path $($args[0]) -Name $($args[1]) | Select-Object -ExpandProperty $($args[1])} -ArgumentList $CustomRegPath, $RegNameLastUpdateDate
$LastRecomposeDate = $RecomposeDate
$PoolType = $HVPoolType
$PoolProvisionType = $HVPoolProvisionType
$PoolAssignType = $HVPoolAssign
$PoolName = $HVPoolNAME

#______________________________________________________________________________________
# Create Function for Testing if vCenter Attribute exists and if not Create the Attribute
function Create-NewCustomAttributes ($CustomAttributeName) {
$error.clear()
Try {
Get-CustomAttribute -Name $CustomAttributeName -ErrorAction SilentlyContinue }
catch { "Get-CustomAttribute *" }
if ($error) {New-CustomAttribute -Name $CustomAttributeName -TargetType VirtualMachine}
}

#______________________________________________________________________________________
# Run the New Custom Attributes Function
Create-NewCustomAttributes $RegNameCorporateBuildVersion
Create-NewCustomAttributes $RegNameWindowsVersion
Create-NewCustomAttributes $RegNameWindowsBuildNumber
Create-NewCustomAttributes $RegNameWindowsRevision
Create-NewCustomAttributes $RegNameLastUpdateDate
Create-NewCustomAttributes $RegNameLastRecomposeDate
Create-NewCustomAttributes $RegNamePoolType
Create-NewCustomAttributes $RegNamePoolProvisionType
Create-NewCustomAttributes $RegNamePoolAssignmentType
Create-NewCustomAttributes $RegNamePoolName
Write-Host "Created Custom Attributes in vCenters"

#______________________________________________________________________________________
# Set vCenter Custom Attributes
Set-Annotation -Entity $VMLine -CustomAttribute $RegNameCorporateBuildVersion -Value $CorporateBuildVersion
Set-Annotation -Entity $VMLine -CustomAttribute $RegNameWindowsVersion -Value $WindowsVersion
Set-Annotation -Entity $VMLine -CustomAttribute $RegNameWindowsBuildNumber -Value $WindowsBuildNumber
Set-Annotation -Entity $VMLine -CustomAttribute $RegNameWindowsRevision -Value $WindowsRevision
Set-Annotation -Entity $VMLine -CustomAttribute $RegNameLastUpdateDate -Value $LastUpdateDate
Set-Annotation -Entity $VMLine -CustomAttribute $RegNameLastRecomposeDate -Value $LastRecomposeDate 
Set-Annotation -Entity $VMLine -CustomAttribute $RegNamePoolType -Value $PoolType
Set-Annotation -Entity $VMLine -CustomAttribute $RegNamePoolProvisionType -Value $PoolProvisionType
Set-Annotation -Entity $VMLine -CustomAttribute $RegNamePoolAssignmentType -Value $PoolAssignType
Set-Annotation -Entity $VMLine -CustomAttribute $RegNamePoolName -Value $PoolName
Write-Host "Set Custom Attributes in vCenters"

#______________________________________________________________________________________
# Set Master VM Cutom Registry Keys
Invoke-Command -ComputerName $VMLineDomain -ScriptBlock {Get-Item -path $($args[0]) | New-ItemProperty -Name  $($args[1]) -Value $($args[2]) -Force | Out-Null} -ArgumentList $CustomRegPath, $RegNameLastRecomposeDate, $LastRecomposeDate
Invoke-Command -ComputerName $VMLineDomain -ScriptBlock {Get-Item -path $($args[0]) | New-ItemProperty -Name  $($args[1]) -Value $($args[2]) -Force | Out-Null} -ArgumentList $CustomRegPath, $RegNamePoolType, $PoolType
Invoke-Command -ComputerName $VMLineDomain -ScriptBlock {Get-Item -path $($args[0]) | New-ItemProperty -Name  $($args[1]) -Value $($args[2]) -Force | Out-Null} -ArgumentList $CustomRegPath, $RegNamePoolProvisionType, $PoolProvisionType
Invoke-Command -ComputerName $VMLineDomain -ScriptBlock {Get-Item -path $($args[0]) | New-ItemProperty -Name  $($args[1]) -Value $($args[2]) -Force | Out-Null} -ArgumentList $CustomRegPath, $RegNamePoolAssignmentType, $PoolAssignType
Invoke-Command -ComputerName $VMLineDomain -ScriptBlock {Get-Item -path $($args[0]) | New-ItemProperty -Name  $($args[1]) -Value $($args[2]) -Force | Out-Null} -ArgumentList $CustomRegPath, $RegNamePoolName, $PoolName
Write-Host "Created Custom Reg Keys on Master Image $VMlineDomain."

#______________________________________________________________________________________
# Set vCenter Notes for the VM.
$Notes = "$RegNameCorporateBuildVersion : $CorporateBuildVersion
$RegNameWindowsVersion : $WindowsVersion
$RegNameWindowsBuildNumber : $WindowsBuildNumber
$RegNameWindowsRevision : $WindowsRevision
$RegNameLastUpdateDate : $LastUpdateDate
$RegNameLastRecomposeDate : $LastRecomposeDate
$RegNamePoolType : $PoolType
$RegNamePoolProvisionType : $PoolProvisionType
$RegNamePoolAssignmentType : $PoolAssignType
$RegNamePoolName : $PoolName"
Write-Host "Created the vCenterNotes Varrible and Populated it with info"

#__________________________________________________________________________
#Set VM nots for each of the Master VMs
set-vm $VMLine -Notes $Notes -Confirm:$false
Write-Host "Set VM Nots for the Master Image VM $VMLine"
}

#______________________________________________________________________________________
#Start the Shutdown Scriopt on Master VM
Write-Host "About to kick off the Shutdown Script on VDI Master $VMLine"
Invoke-Command -ComputerName $VMLineDomain -filepath "$ScriptLocation\$ShutdownScript"
Write-Host "Finished running the Shutdown Script on VDI Master $VMLine"

#______________________________________________________________________________________
#Check to VM powerstate and wait for VM to power off before continuing down the script
$VMPower = get-vm $VMLine
Write-Host "Waiting for VDI Master Image $VMLine to Shutdown"
Do {
    Start-Sleep 30
    $VMPower = get-vm $VMLine
    Write-Host $VMPower.PowerState
    } While ($VMpower.PowerState -eq "PoweredOn")

#______________________________________________________________________________________
#Create SNAPshots for each of the VMs
get-vm $VMLine | Where-Object {$_.powerstate -eq "poweredoff"} | New-Snapshot -Name $ScriptDate -Description "Automated Install of Updates for $ScriptDate" –RunAsync
Write-Host "Created Snapshot for $VMLine Named $ScriptDate Details Below"
#set-vm $VMLine -Notes "Last Recomposed: $ScriptDate" -Confirm:$false
Write-Host "End of Snapshot Config for $VMline"

#______________________________________________________________________________________
#Remove old SNAPshots from each of the VMs
Write-Host "Removing old Snapshots listed below"
get-vm $VMLine | Where-Object {$_.powerstate -eq "poweredoff"} | Get-Snapshot | Where-Object {$_.Created -lt (Get-Date).AddDays($SnapDays)}
get-vm $VMLine | Where-Object {$_.powerstate -eq "poweredoff"} | Get-Snapshot | Where-Object {$_.Created -lt (Get-Date).AddDays($SnapDays)} | Remove-Snapshot -Confirm:$false
Write-Host "Removed old Snapshots from above"

#______________________________________________________________________________________
#PowerOn each of the master VMs
get-vm $VMLine | Where-Object {$_.powerstate -eq "poweredoff"} | Start-VM
Start-Sleep 240
write-host $vmline.name " Is in the powerstate of " $vmline.PowerState

#______________________________________________________________________________________
#Start the SCCM Service on Master VM
Invoke-Command -ComputerName $VMLineDomain -ScriptBlock ${Function:Start-VDIservices} -ArgumentList CCMexec
Write-Host "Start SCCM Service and current status is below"
Invoke-Command -ComputerName $VMLineDomain -ScriptBlock {Get-Service CCMexec}

#______________________________________________________________________________________
#Start the Windows Update Service on Master VM
Invoke-Command -ComputerName $VMLineDomain -ScriptBlock ${Function:Start-VDIservices} -ArgumentList wuauserv
Write-Host "Start Windows Update Service and current status is below"
Invoke-Command -ComputerName $VMLineDomain -ScriptBlock {Get-Service wuauserv}

#______________________________________________________________________________________
#Copy Log from from the VDI Master VM to share location
Copy-Item -Path $MasterLogLocation\$ScriptDate\ -Destination $ShareLogLocation\ -Recurse -Force
Write-host "Copy Log files from VDI Master VM $VMLine to Share Location $ShareLogLocation\$ScriptDate"

#______________________________________________________________________________________
#Delete Log folder and all contents on Master VM
Remove-Item -Path $MasterLogLocation\$ScriptDate -Force -Recurse
Write-Host "Removed log folder from VDI Master VM $VMLine"

#______________________________________________________________________________________
#Update Change With with Log files from from Install and Shutdown Script
if($RunServiceNow -eq '1')
{
    #______________________________________________________________________________________
    #Update Change Notes with Installed Software Before Report
    $UpdateSNChangeNotes = [IO.File]::ReadAllText("$ShareLogLocation\$ScriptDate\InstalledSoftwareBefore.$VMline.txt")
    Update-Change $UpdateSNChangeNotes
    Write-host "Added Contents of $ShareLogLocation\$ScriptDate\InstalledSoftwareBefore.$VMline.txt to Service Now Change $SNChangeID."

    #______________________________________________________________________________________
    #Update Change Notes with Installed Software After Report
    $UpdateSNChangeNotes = [IO.File]::ReadAllText("$ShareLogLocation\$ScriptDate\InstalledSoftwareAfter.$VMline.txt")
    Update-Change $UpdateSNChangeNotes
    Write-host "Added Contents of $ShareLogLocation\$ScriptDate\InstalledSoftwareAfter.$VMline.txt to Service Now Change $SNChangeID."

    #______________________________________________________________________________________
    #Update Change Notes with Windows Updates Installed Report
    $UpdateSNChangeNotes = [IO.File]::ReadAllText("$ShareLogLocation\$ScriptDate\WindowsUpdates.$VMline.txt")
    Update-Change $UpdateSNChangeNotes
    Write-host "Added Contents of $ShareLogLocation\$ScriptDate\WindowsUpdates.$VMline.txt to Service Now Change $SNChangeID."
   
    #______________________________________________________________________________________
    #Update Change Notes with Optimization Error Report Report
    $UpdateSNChangeNotes = [IO.File]::ReadAllText("$ShareLogLocation\$ScriptDate\OptimizationErrorLog.$VMline.txt")
    Update-Change $UpdateSNChangeNotes
    Write-host "Added Contents of $ShareLogLocation\$ScriptDate\OptimizationErrorLog.$VMline.txt to Service Now Change $SNChangeID."

}


}

#______________________________________________________________________________________
#Connect to Connection servers
foreach ($HVServer in $HVServers) 
{
    Connect-HVServer $HVServer -Credential $VMwareSVCCreds
    Write-Host "Connected to $HVServer"

    #______________________________________________________________________________________
    #Get Pool Names
    $poolsupdate = Get-HVPool
    foreach($poolID in $poolsupdate) 
    {
        $HVPoolIDs = $poolID | get-hvpoolspec | ConvertFrom-Json
        $HVPoolIDText = $HVPoolIDs.base.name
        
        #______________________________________________________________________________________
        # Set recompose delay based on if pool is Test or Prod
        if($HVPoolIDText -like $TestPoolNC)
            {
                $RecomposeDate = (Get-Date).AddHours($RCTimeDelay1)
                Write-host "Pool Named"$HVPoolIDs.base.name"is a Test pool and will be Recoposed or Push Image at $RecomposeDate"
            }else {$RecomposeDate = (Get-Date).AddHours($RCTimeDelay2)}
            Write-host "Pool Named"$HVPoolIDs.base.name"is a Production pool and will be Recoposed or Push Image at $RecomposeDate"

        #______________________________________________________________________________________
        #Does Image Push to Instant Clone Pool    
        if($HVPoolIDs.AutomatedDesktopSpec.provisioningType -like 'INSTANT_CLONE_ENGINE')
            {
                Start-HVPool -SchedulePushImage -Pool $HVPoolIDs.base.name -LogoffSetting WAIT_FOR_LOGOFF -ParentVM $HVPoolIDs.AutomatedDesktopSpec.virtualCenterProvisioningSettings.VirtualCenterProvisioningData.parentVm -SnapshotVM $ScriptDate -StartTime $RecomposeDate
                Write-Host "Recomposing the Pool"$HvpoolIDs.Base.name"with ParentVM"$HvpoolIDs.AutomatedDesktopSpec.virtualCenterProvisioningSettings.VirtualCenterProvisioningData.parentVm"with pool type of"$HVPoolIDs.AutomatedDesktopSpec.provisioningType"at"$RecomposeDate" with the snapshot of "$ScriptDate"."
            }
        
        #______________________________________________________________________________________
        #Does Recompose to Linked Clone Pool
        elseif($HVPoolIDs.AutomatedDesktopSpec.provisioningType -like 'VIEW_COMPOSER')
            {
                Start-HVPool -Recompose -Pool $HVPoolIDs.base.name -LogoffSetting WAIT_FOR_LOGOFF -ParentVM $HVPoolIDs.AutomatedDesktopSpec.virtualCenterProvisioningSettings.VirtualCenterProvisioningData.parentVm -SnapshotVM $ScriptDate -StartTime $RecomposeDate
                Write-Host "Recomposing the Pool"$HvpoolIDs.Base.name"with ParentVM"$HvpoolIDs.AutomatedDesktopSpec.virtualCenterProvisioningSettings.VirtualCenterProvisioningData.parentVm"with pool type of"$HVPoolIDs.AutomatedDesktopSpec.provisioningType"at"$RecomposeDate" with the snapshot of "$ScriptDate"."
            }
       }
    #______________________________________________________________________________________
    #Disconnect from Connection Server
    Disconnect-HVServer $hvserver -Confirm:$false
    Write-Host "Disconnected from $hvServer"
}

#______________________________________________________________________________________
# Stop Logging
Stop-Transcript

$UpdateSNChangeNotes = "Completed Reresh of all Pools and scheduled the Recompose and Image Push"
$UpdateSNChangeUpdateChangeSummary = [IO.File]::ReadAllText("$ShareLogLocation\$ScriptDate\WeeklyVDIUpdates.txt")
Update-ChangeClosePrep $UpdateSNChangeNotes $UpdateSNChangeUpdateChangeSummary
Write-host "Updated Notes and Change Summary to Close out Service Now Change $SNChangeID."

$UpdateSNIncidentNotes = "Completed all tasks in the Change $SNChangeID"
$UpdateSNIncidentCloseNotes = "Completed the Change $SNChangeID"
Update-IncidentClosePrep $UpdateSNIncidentNotes $UpdateSNIncidentCloseNotes
Write-host "Update Notes and Close Notes to Close out Service Now Incident $SNIncidentID."

Update-ChangeClose
Update-IncidentClose