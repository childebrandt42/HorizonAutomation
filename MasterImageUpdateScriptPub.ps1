# Horizon Master Image Automated Update Process
# Chris Hildebrandt
# 12-5-2018
# Version 2.0 updated 12-5-2018
# Script will power on all master, Run software updates, and run the logoff Script, Snapshot the Masters, Update the Notes, Recompose or Push Images  

#______________________________________________________________________________________
#User Editable Varibles

$vCenterServers = @("ServerName.Domain.com","ServerName.Domain.com")  #Enter vCenters FQDN in format of ("vCenter1","vCenter2")  
$HVServers = @("ServerName.Domain.com","ServerName.Domain.com")   #Enter HVServer FQDN in format of ("HVServer1","HVServer2")   
$MSTRfolder = "Master Images"   #vCenter Folder name where the master images are located. It will Update all VMs in that folder so only put current master images in that folder.
$SnapDays = "-30"   #Number of Days to keep past snapshots on the VDI Master VM must be a negitive number
$RCTimeDelay1 = '1' #Recompose Test Pools Delay in hours
$RCTimeDelay2 = '48'    #Recompose Prod Pools Delay in hours
$TestPoolNC = "Test"    #Test pool naming convention that diffirentiates them from standard pool.

$LogLocation = "C:\VDI_Tools\Logs"  #Location to save the logs on the Master VM. They will be copied to share and deleted upon completion of the script.
$ScriptLocation = "C:\VDI_Tools\Scripts"    #Location for the Scripts on the Master VMs. This folder will remain after script completes
$CloneToolsLocation = "C:\VDI_Tools\CloneTools" #Location on the Master VMs where the VMware Optimizer is stored and any other tools.
$SleepTimeInS = "180"   #Wait timer in minuits for windows to install Windows Updates via Windows Update Service. NOT needed if using SCCM to install updates. 
$AdobeUpdate = "AdobeAcrobatUpdate.ps1" #Name of Adobe Update Script. NOT needed if using SCCM to install updates.
$FlashUpdate = "FlashUpdate.ps1"    #Name of Flash Update Script. NOT needed if using SCCM to install updates.
$FirefoxUpdate = "FireFoxUpdate.ps1"    #Name of FireFox Update Script. NOT needed if using SCCM to install updates.
$JavaUpdate = "JavaUpdate.ps1"  #Name of Java Update Script. NOT needed if using SCCM to install updates.
$CustomRegPath = "HKLM:\Software\YourCompanyName\Horizon\" #Custom Registry Path to create a Key that keeps record of when it was last updated.
$RunOptimizer = '1' #To run Optimizer enter 1, If you do not want to run enter 0
$RunSEP = '1' #To run Optimizer enter 1, If you do not want to run enter 0
$OptimizerTemplateNamingPrefix = "TemplateName" #Template Naming i.e. "CompanyTemplate10" for windows 10 Please make sure to append the namy by 10 for win10, 8 for win8, and 7 for win7
$VMwareOptimizerName = "VMwareOSOptimizationTool.exe"   #Name Of the VMware Optimizer Tool
$ShareLogLocation = "\\ShareName.domain.com\VDI_Tools\Logs" #Network Share to save the logs. This will be the collection point for all logs as each VDI Master VM will copy there local logs to this location.
$ShareScriptLocation = "\\ShareName.domain.com\VDI_Tools\Scripts"   #Network Share location where all the scripts are stored and distributed from. 
$ShareCloneToolsLocation = "\\ShareName.domain.com\VDI_Tools\CloneTools"    #Network Share Location where all the Clone Tools are stored and distributed from.

#______________________________________________________________________________________
# Do not Edit Below! Do not Edit Below! Do not Edit Below! Do not Edit Below!
#______________________________________________________________________________________

#______________________________________________________________________________________
#Enviroment Varribles Do not EDIT
$ScriptDate = Get-Date -Format MM-dd-yyyy

#______________________________________________________________________________________
#Start the Debug Logging
$ErrorActionPreference="SilentlyContinue"
Stop-Transcript | out-null
$ErrorActionPreference = "Continue"
Start-Transcript -Force -Path $LogLocation\$ScriptDate\WeeklyVDIUpdates.txt
Write-Host "Start Debug Logs"

#______________________________________________________________________________________
#Import VMware Modules
Install-Module -Name VMware.PowerCLI
Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Force -confirm:$false
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
# Capture installed Software before updating
`$InstalledSoftwareAfter = get-software | Format-Table | Out-File "$LogLocation\$ScriptDate\InstalledSoftwareAfter.`$MasterVMHostName.txt"
write-host "Capture Installed Software After update and save to Logs Folder"

#______________________________________________________________________________________
# Check if Custom REG key is created if not create LastRecomposeDate REG key with 1/1/2000 date
if ( -not (Test-Path -Path $CustomRegPath)) {New-Item -path $CustomRegPath -Force | New-ItemProperty -Name  LastRecomposeDate -Value "01/01/2000" -Force | Out-Null}
write-host "Check if Custom key has been created and if not create registry key and set value to 1-1-2000"

#______________________________________________________________________________________
# Get LastRecomposeData value
`$LastRecomposeDate = Get-ItemProperty -Path $CustomRegPath -Name LastRecomposeDate | Select-Object -ExpandProperty LastRecomposeDate
write-host "Get Last recompose date from custom Registry key, Last Recompose date was `$LastRecomposeDate"

#______________________________________________________________________________________
# Get Windows Udates installed since last recompose. 
`$WindowsUpdates = Get-HotFix | Where-Object -Property InstalledOn -GT `$LastRecomposeDate | Out-File -FilePath "$LogLocation\$ScriptDate\WindowsUpdates.`$MasterVMHostName.txt"
write-host "Create TXT file with Installed updates as of `$LastRecomposeDate"

#______________________________________________________________________________________
#Set REG Key to store Last Recompose Date
Set-ItemProperty -Path $CustomRegPath -Name LastRecomposeDate -Value "`$ScriptDate2" -Force
write-host "Set Last Recompose date to `$ScriptDate2"

"@  |  Set-Content "$ScriptLocation\$InstallUpdatesScript"
Write-Host "Created Install Script and saved it to location $ScriptLocation\$InstallUpdatesScript"

#______________________________________________________________________________________
#Define ShutdownScript File Name Var
$ShutdownScript = "ShutDownScript.ps1"

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
#Connect to vCenters
foreach ($vCenterServer in $vCenterServers) 
{
    Connect-VIServer $vCenterServer -Credential $VMwareSVCCreds
    Write-Host "Connected to vCenter Server $vCenterServer"
}

#______________________________________________________________________________________
#Get VM lists List1 for Test Pools and List2 is Poduction Pools
$MSTRVMList = Get-Folder $MSTRfolder | get-vm
Write-host "Build list of Master VM's to Update and Recompose or Push the Pools. VMs to complete are "
$MSTRVMList.Name | Format-list

Write-Host "About to Invoke Update"


#______________________________________________________________________________________
#Create Install and Shutdown Scripts
#______________________________________________________________________________________


#______________________________________________________________________________________
#Run Remote Scripts on each of the VDI Master Imagess
foreach ($VMLine in $MSTRVMList)
{

#______________________________________________________________________________________
#Define VDI Master VMs Log Location
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
Invoke-Command -ComputerName $VMLine -filepath "$ScriptLocation\$InstallUpdatesScript"
Write-Host "Finished running the Install Updates Script on VDI Master $VMLine"

#______________________________________________________________________________________
#Restart the Master VM
Write-Host "Restarting VDI Master VM $VMLine"
Restart-Computer -ComputerName $VMLine -Force -Wait

#______________________________________________________________________________________
#ReStart the Debug Logging
$ErrorActionPreference="SilentlyContinue"
Stop-Transcript | out-null
$ErrorActionPreference = "Continue"
Start-Transcript -Force -Path $LogLocation\$ScriptDate\WeeklyVDIUpdates.txt -Append
Write-Host "Start Debug Logs"


#______________________________________________________________________________________
#Start the Shutdown Scriopt on Master VM
Write-Host "About to kick off the Shutdown Script on VDI Master $VMLine"
Invoke-Command -ComputerName $VMLine -filepath "$ScriptLocation\$ShutdownScript"
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
set-vm $VMLine -Notes "Last Recomposed: $ScriptDate" -Confirm:$false
Write-Host "End of Snapshot Config for $VMline"

#______________________________________________________________________________________
#Set VM nots for each of the Master VMs
set-vm $VMLine -Notes "Last Recomposed: $ScriptDate" -Confirm:$false
Write-Host "Set VM Nots for Last Recompose Date: $ScriptDate"
get-vm $VMLine | Select-Object Name,Notes

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
write-host Get-vm $vmline.name " Is in the powerstate of " Get-vm $vmline.PowerState

#______________________________________________________________________________________
#Start the SCCM Service on Master VM
Invoke-Command -ComputerName $VMLine -ScriptBlock ${Function:Start-VDIservices} -ArgumentList CCMexec
Write-Host "Start SCCM Service and current status is below"
Invoke-Command -ComputerName $VMLine -ScriptBlock {Get-Service CCMexec}

#______________________________________________________________________________________
#Start the Windows Update Service on Master VM
Invoke-Command -ComputerName $vmline -ScriptBlock ${Function:Start-VDIservices} -ArgumentList wuauserv
Write-Host "Start Windows Update Service and current status is below"
Invoke-Command -ComputerName $VMLine -ScriptBlock {Get-Service wuauserv}

#______________________________________________________________________________________
#Copy Log from from the VDI Master VM to share location
Copy-Item -Path $MasterLogLocation\$ScriptDate\ -Destination $ShareLogLocation\ -Recurse -Force
Write-host "Copy Log files from VDI Master VM $VMLine to Share Location $ShareLogLocation\$ScriptDate"

#______________________________________________________________________________________
#Delete Log folder and all contents on Master VM
Remove-Item -Path $MasterLogLocation\$ScriptDate -Force -Recurse
Write-Host "Removed log folder from VDI Master VM $VMLine"
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
        
        #______________________________________________________________________________________
        # Set recompose delay based on if pool is Test or Prod
        if($HVPoolIDs.base.name -like $TestPoolNC)
            {
                $RecomposeDate = (Get-Date).AddHours($RCTimeDelay1)
                Write-host "Pool Named $HVPoolIDs.base.name is a Test pool and will be Recoposed or Push Image at $RecomposeDate"
            }else {$RecomposeDate = (Get-Date).AddHours($RCTimeDelay2)}
            Write-host "Pool Named $HVPoolIDs.base.name is a Production pool and will be Recoposed or Push Image at $RecomposeDate"

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