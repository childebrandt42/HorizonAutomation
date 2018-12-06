# Shutdown Script to Prepare Master Image for Recompose or Image Push
# Chris Hildebrandt
# 10-18-2018
# Ver 1.5
# Script will Run Disk Cleanup, Defrag, Run SEP scan force update, Run Optimizer, and shutdown services, clear DNS and IP, and Shutdown VM

#______________________________________________________________________________________
# Check Version of OS
$winVersion = "0"
$OSVersion = [System.Environment]::OSVersion.Version
if ($OSversion.Major -eq "10" -and $OSversion.Minor -eq "0"){ $winVersion = "10" }
elseif ($OSversion.Major -eq "6" -and $OSversion.Minor -eq "2") { $winVersion = "8" }
elseif ($OSversion.Major -eq "6" -and $OSversion.Minor -eq "1") { $winVersion = "7" }
Write-host "Windows Version is $winVersion"

#______________________________________________________________________________________
# Varibles
$ScriptDate = Get-Date -Format MM-dd-yyyy
$FQDN = (Get-WmiObject win32_computersystem).DNSHostName+"."+(Get-WmiObject win32_computersystem).Domain
$LoginFile = "C:\VDI_Tools\Logs\$ScriptDate\$FQDN\ShutdownScript_$ScriptDate.txt"
$defragbat = "C:\VDI_Tools\Scripts\dfrag.bat"       #Temporary Location where it saves the Defrag.bat file.

#______________________________________________________________________________________
# Optimizer Varibles
$RunOptimizer = 1   # To run Optimizer enter 1, If you do not want to run enter 0
$OptimizerLocation = "C:\VDI_Tools\CloneTools\VMwareOSOptimizationTool"
$OptimizerTemplate = "C:\VDI_Tools\CloneTools\ARDx_Windows${winVersion}.xml"
$OptimizerLogLocation = "C:\VDI_Tools\Logs\$ScriptDate\$FQDN\"
$OptimizerErrorLogLocation = "C:\VDI_Tools\Logs\$ScriptDate\$FQDN\OptimizationErrorLog.txt"
$OptimizerbatFile = "C:\VDI_Tools\Scripts\Optimizer.$fqdn.bat"         #Temporary Location where it saves the Optimizer.bat file.
$RunSEP = 1 # To run Optimizer enter 1, If you do not want to run enter 0

#______________________________________________________________________________________
# Start the Debug Logging
$ErrorActionPreference="SilentlyContinue"
Stop-Transcript | out-null
$ErrorActionPreference = "Continue"
Start-Transcript -path $LoginFile

Write-host "Windows Version is $winVersion"

#______________________________________________________________________________________
# Define functions

#______________________________________________________________________________________
# Start Services Function
function Start-VDIservices ($SVCname) {
    Set-Service -Name $svcname -StartupType Automatic
    Start-Service -Name $svcname
    Do {
    $svc = Get-Service -Name $svcname
    Start-Sleep 2
    } While ( $svc.Status -ne "Running" )
    }

#______________________________________________________________________________________
# Stop Services Function
function Stop-VDIservices ($SVCname) {
    Stop-Service -Name $svcname
    Do {
    $svc = Get-Service -Name $svcname
    Start-Sleep 2
    } While ( $svc.Status -ne "Stopped" )
    Set-Service -Name $svcname -StartupType disabled
    }

#______________________________________________________________________________________
# Do while proc Like Function
function DoWhile-LikeProcName ($Process) {
    Do
    {
    "Processes $Process is still running"
    $proc = Get-Process
    start-sleep 10
    } While ($proc.name -like "*$Process*")
    }

#______________________________________________________________________________________
# Do while Proc Equals Function
function DoWhile-EQProcName ($Process) {
    Do
    {
    "Processes $Process is still running"
    $proc = Get-Process
    start-sleep 10
    } While ($proc.name -eq "$Process")
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
$DfragStatus = get-service -name defragsvc | Select-Object Name,Status,Starttype
if($DfragStatus.Status -eq 'Stopped')
{
#______________________________________________________________________________________
# Start Defrag Service
Start-VDIservices defragsvc
Write-host "Start Defrag Service"
}

#______________________________________________________________________________________
# Create Temp Defrag BAT
"defrag c: /U /V" | Set-Content $defragbat
Write-host "Create Defrag Temp .bat file"

#______________________________________________________________________________________
# Start Defrag BAT
Start-Process $defragbat
Write-host "Running Defrag Cleanup"

#______________________________________________________________________________________
# Check status of Disk Defrag
DoWhile-EQProcName Defrag

#______________________________________________________________________________________
# Remove Temp Defrag BAT
Remove-Item -Path $defragbat
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
if($RunSEP -eq '1')
{
    Write-host "Starting the SEP Update and Cleanup Process"
#______________________________________________________________________________________
# Check to see if SEP is installed and prep for Recompose
#______________________________________________________________________________________
# If Check for SEP install
$CheckSEP = get-service -name SepMasterService | Select-Object Name,Status,Starttype
if($CheckSEP.Name -eq 'SepMasterService')
    {

    #______________________________________________________________________________________
    # Force SEP checking and Update
    Start-Process -Filepath "C:\Program Files (x86)\Symantec\Symantec Endpoint Protection\SepLiveUpdate.exe"
    Write-host "Liveupdate SEP Config"

    #______________________________________________________________________________________ 
    # Check status of SEP Update
    DoWhile-LikeProcName SepLiveUpdate

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
    $SEPStatus = get-service -name SepMasterService | Select-Object Name,Status,Starttype
    if($SEPStatus.Status -eq 'Running')
    {
        #______________________________________________________________________________________
        # Stop SEP Service and Prep for Clone
        Start-Process -Filepath "C:\Program Files (x86)\Symantec\Symantec Endpoint Protection\smc.exe" -argumentlist '-stop'
        Write-host "Stopped SEP Service"
    }

        $filesToClean = @("C:\sephwid.xml","C:\communicator.dat","C:\Program Files\Common Files\Symantec Shared\HWID\sephwid.xml","C:\Program Files\Common Files\Symantec Shared\HWID\communicator.dat", "C:\Windows\Temp\sephwid.xml","C:\Windows\Temp\communicator.dat","C:\Documents and Settings\*\Local Settings\Temp\sephwid.xml","CC:\Documents and Settings\*\Local Settings\Temp\communicator.dat","C:\Users\*\AppData\Local\Temp\sephwid.xml","C:\Users\*\AppData\Local\Temp\communicator.dat")

        foreach($file in $filesToClean)
        {
            if(Test-Path $file) {Remove-Item -Path $file -Recurse -force}
            Write-Host "Removed File $File"
        }
        
        #______________________________________________________________________________________ 
        # Remove SEP Reg Keys

        #______________________________________________________________________________________ 
        # Remove 32bit SEP Reg Keys
        if (Test-Path -Path "HKLM:\SOFTWARE\Symantec\Symantec Endpoint Protection\SMC\SYLINK\SyLink")
        {
            #Remove-ItemProperty -Path "HKLM:\SOFTWARE\Symantec\Symantec Endpoint Protection\SMC\SYLINK\SyLink" -Name "ForceHardwareKey"
            Remove-ItemProperty -Path "HKLM:\SOFTWARE\Symantec\Symantec Endpoint Protection\SMC\SYLINK\SyLink" -Name "HardwareID"
            Remove-ItemProperty -Path "HKLM:\SOFTWARE\Symantec\Symantec Endpoint Protection\SMC\SYLINK\SyLink" -Name "HostGUID"
            Write-host "Removed SEP 32 Bit Reg Keys"
        }
        if (Test-Path -Path "HKLM:\SOFTWARE\WOW6432Node\Symantec\Symantec Endpoint Protection\SMC\SYLINK\SyLink")
        {
            #Remove-ItemProperty -Path "HKLM:\SOFTWARE\WOW6432Node\Symantec\Symantec Endpoint Protection\SMC\SYLINK\SyLink" -Name "ForceHardwareKey"
            Remove-ItemProperty -Path "HKLM:\SOFTWARE\WOW6432Node\Symantec\Symantec Endpoint Protection\SMC\SYLINK\SyLink" -Name "HardwareID"
            Remove-ItemProperty -Path "HKLM:\SOFTWARE\WOW6432Node\Symantec\Symantec Endpoint Protection\SMC\SYLINK\SyLink" -Name "HostGUID"
            Write-host "Removed SEP 64 Bit Reg Keys"
         }
    
    }
    Write-host "Finished SEP Cleanup and Prep for Cloneing"
}

#______________________________________________________________________________________ 
# Run VMware optimizer
if($RunOptimizer -eq '1')
{
    Write-host "Prepairing to run VMware Optimizer"

#______________________________________________________________________________________ 
# Create Temp Bat file to optimize OS
@"
start /wait $OptimizerLocation -t $OptimizerTemplate -r $OptimizerLogLocation
start /wait $OptimizerLocation -o recommended -t $OptimizerTemplate -v > $OptimizerErrorLogLocation 2>&1
start /wait $OptimizerLocation -t $OptimizerTemplate -r $OptimizerLogLocation
"@ | Set-Content $OptimizerbatFile

Write-host "Built Temporary VMware Optimizer .BAT file based on Template $OptimizerTemplate"
    #______________________________________________________________________________________ 
    # Run BAT file to optimize OS
    Start-Process $OptimizerbatFile
    Write-host "Running VMware Optimizer .BAT based on Template $OptimizerTemplate"

    #______________________________________________________________________________________ 
    # Check status of the Optimize and pause
    DoWhile-EQProcName CMD

    #______________________________________________________________________________________ 
    # Remove Temp BAT file
    Remove-Item -Path $OptimizerbatFile
    Write-host "Removing Temporary VMware Optimizer .BAT file"

}
#______________________________________________________________________________________
# Check if VM is managed by SCCM
$CheckSCCM = get-wmiobject win32_Service | Where-Object {$_.Name -eq "CCMexec"} | Select-Object Name
Write-host "Check if SCCM service is installed"

#______________________________________________________________________________________
# If Else statement based on if VM is managed by SCCM
if($CheckSCCM.Name -eq 'CCMexec')
{
    Write-host "VM is managed by SCCM"
    #______________________________________________________________________________________
    # Check status of SCCM Service.
    $SCCMStatus = get-service -name "CCMexec" | Select-Object Status
    Write-host "Check if SCCM service CCMexec is running"
    if($SCCMStatus.Status -eq 'Running')
        {   
            Stop-VDIservices CCMexec
            Write-host "Stop SCCM Service CCMexec"
        }
    
    #______________________________________________________________________________________
    # Check status of SCCM Service if stopped start.
    $AdaptivaStatus = get-service -name "adaptivaclient" | Select-Object Status
    Write-host "Check if Adaptiva Client is installed"
    if($AdaptivaStatus.Status -eq 'Running')
        {
            Stop-VDIservices adaptivaclient
            Write-host "Stop Adaptiva Services"
        }
}

#______________________________________________________________________________________
# Check status of Windows Update Service.
$WindowsUpdatateStatus = get-service -name wuauserv | Select-Object Name,Status,Starttype
if($WindowsUpdatateStatus.Status -eq 'Running')
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
# Check status of VSS Service if stopped start.
$DfragStatus = get-service -name VSS | Select-Object Name,Status,Starttype
if($DfragStatus.Status -eq 'Running')
{
    #______________________________________________________________________________________
    # Clean out Volume Shadow Copys
    .\vssadmin.exe delete shadows /all
    Write-host "Clean out Volume Shadow Copys"
}

#______________________________________________________________________________________
# Clean out Windows PreFetch
if (Test-Path "c:\Windows\Prefetch\")
{
Remove-item -Path "c:\Windows\Prefetch\*.*" -Force
Write-host "Remove windows Prefetch files"
}

#______________________________________________________________________________________
# Check status if AppVolumes is installed and stop the service.
$AppVStatus = get-service | Select-Object Name,Status,Starttype
if($AppVStatus.Name -eq 'svservice')
{
    Stop-Service -Name svservice
    Write-host "Stop AppVolumes Service"
}

#______________________________________________________________________________________
# Clear all event logs
wevtutil el | Foreach-Object {wevtutil cl "$_"}
Write-host "Clear Windows Event Logs"

#______________________________________________________________________________________
# Stop Logging
Stop-Transcript

#______________________________________________________________________________________
# Release IP, Flush DNS and Shutdown VM
ipconfig /release
ipconfig /flushdns
shutdown -s
