# Recompose Master Image (Step2)
# Chris Hildebrandt
# 10-19-2018
# Ver 1.5
# Script will Copy Files to Image, Enable SCCM service, Install SCCM updates,Wait X seconds, and Reboot

#______________________________________________________________________________________
#Varibles
$ScriptDate = Get-Date -Format MM-dd-yyyy
$ScriptDate2 = Get-Date -UFormat %m/%d/%y                       #Date used for custom Registry Key and Windows Updates Report
$FQDN = (Get-WmiObject win32_computersystem).DNSHostName+"."+(Get-WmiObject win32_computersystem).Domain
$LogLocation = "C:\VDI_Tools\Logs\$ScriptDate\$FQDN"
$SleepTimeInS = "60"                                            #Sleep time in seconds
$AdobeUpdate = "C:\VDI_Tools\Scripts\AdobeAcrobatUpdate.ps1"    #Adobe Acrobat Update Script
$FlashUpdate = "C:\VDI_Tools\Scripts\FlashUpdate.ps1"           #Flash Update Script
$FirefoxUpdate = "C:\VDI_Tools\Scripts\FireFoxUpdate.ps1"       #Firefox Update Script
$JavaUpdate = "C:\VDI_Tools\Scripts\JavaUpdate.ps1"             #Java Update Script
$ShutdownScript = "C:\VDI_Tools\Scripts\ShutdownScript.ps1"     #Shutdown Script Location
$CustomRegPath = "HKLM:\Software\Company\Horizon\"              #Custom Registry Location. Replace Company with your Company Name
#______________________________________________________________________________________
#Copy Locations
$CopyshareLocation = "\\share\VDI_Tools\"       #Remote Share location
$DestinationLocation = "C:\VDI_Tools\"

#______________________________________________________________________________________
#Start the Debug Logging
$ErrorActionPreference="SilentlyContinue"
Stop-Transcript | out-null
$ErrorActionPreference = "Continue"
Start-Transcript -path "$loglocation\RecomposeMSTRInstallUpdates_$ScriptDate.txt"

#______________________________________________________________________________________
# Define functions

#______________________________________________________________________________________
# Import Function Get-Software
# Created By Boe prox HTTPS://MCPMAG.COM
Function Get-Software  {
  [OutputType('System.Software.Inventory')]
  [Cmdletbinding()] 
  Param( 
    [Parameter(ValueFromPipeline=$True,ValueFromPipelineByPropertyName=$True)] 
    [String[]]$Computername=$env:COMPUTERNAME
    )         
  Begin {
  }
  Process  {     
  ForEach  ($Computer in  $Computername){ 
  If  (Test-Connection -ComputerName  $Computer -Count  1 -Quiet) {
  $Paths  = @("SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Uninstall","SOFTWARE\\Wow6432node\\Microsoft\\Windows\\CurrentVersion\\Uninstall")         
  ForEach($Path in $Paths) { 
  Write-Verbose  "Checking Path: $Path"

  #  Create an instance of the Registry Object and open the HKLM base key 
  Try  { 
  $reg=[microsoft.win32.registrykey]::OpenRemoteBaseKey('LocalMachine',$Computer,'Registry64') 
  } Catch  { 
  Write-Error $_ 
  Continue 
  } 

  #  Drill down into the Uninstall key using the OpenSubKey Method 
  Try  {
  $regkey=$reg.OpenSubKey($Path)  

  # Retrieve an array of string that contain all the subkey names 
  $subkeys=$regkey.GetSubKeyNames()      

  # Open each Subkey and use GetValue Method to return the required  values for each 
  ForEach ($key in $subkeys){   
  Write-Verbose "Key: $Key"
  $thisKey=$Path+"\\"+$key 
  Try {  
  $thisSubKey=$reg.OpenSubKey($thisKey)   

  # Prevent Objects with empty DisplayName 
  $DisplayName =  $thisSubKey.getValue("DisplayName")
  If ($DisplayName  -AND $DisplayName  -notmatch '^Update  for|rollup|^Security Update|^Service Pack|^HotFix') {
  $Date = $thisSubKey.GetValue('InstallDate')
  If ($Date) {
  Try {
  $Date = [datetime]::ParseExact($Date, 'yyyyMMdd', $Null)
  } Catch{
  Write-Warning "$($Computer): $_ <$($Date)>"
  $Date = $Null
  }
  } 

  # Create New Object with empty Properties 
  $Publisher =  Try {
  $thisSubKey.GetValue('Publisher').Trim()
  } 
  Catch {
  $thisSubKey.GetValue('Publisher')
  }
  $Version = Try {

  #Some weirdness with trailing [char]0 on some strings
  $thisSubKey.GetValue('DisplayVersion').TrimEnd(([char[]](32,0)))
  } 
  Catch {
  $thisSubKey.GetValue('DisplayVersion')
  }
  $UninstallString =  Try {
  $thisSubKey.GetValue('UninstallString').Trim()
  } 
  Catch {
  $thisSubKey.GetValue('UninstallString')
  }
  $InstallLocation =  Try {
  $thisSubKey.GetValue('InstallLocation').Trim()
  } 
  Catch {
  $thisSubKey.GetValue('InstallLocation')
  }
  $InstallSource =  Try {
  $thisSubKey.GetValue('InstallSource').Trim()
  } 
  Catch {
  $thisSubKey.GetValue('InstallSource')
  }
  $HelpLink = Try {
  $thisSubKey.GetValue('HelpLink').Trim()
  } 
  Catch {
  $thisSubKey.GetValue('HelpLink')
  }
  $Object = [pscustomobject]@{
  Computername = $Computer
  DisplayName = $DisplayName
  Version  = $Version
  InstallDate = $Date
  Publisher = $Publisher
  UninstallString = $UninstallString
  InstallLocation = $InstallLocation
  InstallSource  = $InstallSource
  HelpLink = $thisSubKey.GetValue('HelpLink')
  EstimatedSizeMB = [decimal]([math]::Round(($thisSubKey.GetValue('EstimatedSize')*1024)/1MB,2))
  }
  $Object.pstypenames.insert(0,'System.Software.Inventory')
  Write-Output $Object
  }
  } Catch {
  Write-Warning "$Key : $_"
  }   
  }
  } Catch  {}   
  $reg.Close() 
  }                  
  } Else  {
  Write-Error  "$($Computer): unable to reach remote system!"
  }
  } 
  } 
}  

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
#Check is VDI Masters Account Password file exists
if(-Not (Test-Path -Path "C:\VDI_Tools\Scripts\VDIMastersPassword.txt" ))
{
    Write-Host "Credentials file not found proceeding to creation"
    #______________________________________________________________________________________
    #Create Secure Password File
    Get-Credential -Message "Enter VDI Master Admin Account Domain\Username" | Export-Clixml "C:\VDI_Tools\Scripts\VDIMastersPassword.txt"
    write-host "Create Credentials File by prompting for creds"
}

#______________________________________________________________________________________
#Import Secure Creds for use.
$VDIMSTRCreds = Import-Clixml "C:\VDI_Tools\Scripts\VDIMastersPassword.txt"
write-host "Import Secure Creds"

#______________________________________________________________________________________
#Copy Clone Tools and Scripts
Copy-Item $CopyshareLocation -Destination $DestinationLocation -Recurse
write-host "Copy Files from Share"

#______________________________________________________________________________________
# Capture installed Software before updating
get-software | Format-Table | Out-File "$LogLocation\InstalledSoftwareBefore.txt"
write-host "Capture Installed Software Prior to update and save to Logs Folder"

#______________________________________________________________________________________
#Check to see if machine is managed by SCCM
$CheckSCCM = get-wmiobject win32_Service | Where-Object {$_.Name -eq "CCMexec"} | Select-Object Name
write-host "Check if VM is managed by SCCM"

#______________________________________________________________________________________
#If Else statement based on if VM is managed by SCCM
if($CheckSCCM.Name -eq 'CCMexec')
{
    write-host "VM is managed by SCCM"
    #______________________________________________________________________________________
    #Check Status of SCCM service
    $SCCMStatus = get-service -name CCMexec | Select-Object Name,Status,Starttype
    if($SCCMStatus.Status -eq 'Stopped')
    {   
        #______________________________________________________________________________________
        #Set SCCM to Automatic and Start Service
        Start-VDIservices CCMexec
        write-host "Start SCCM Services"
    }

    #______________________________________________________________________________________
    #Start SCCM Patching
    $AppEvalState0 = "0" 
    $AppEvalState1 = "1" 
    $Application = (Get-WmiObject -Namespace "root\ccm\clientSDK" -Class CCM_SoftwareUpdate | Where-Object { $_.EvaluationState -like "*$($AppEvalState0)*" -or $_.EvaluationState -like "*$($AppEvalState1)*"})
    Invoke-WmiMethod -Class CCM_SoftwareUpdatesManager -Name InstallUpdates -ArgumentList (,$Application) -Namespace root\ccm\clientsdk
    write-host "Invoke SCCM install of Updates"

    #______________________________________________________________________________________
    #Pause Script to wait for Updates to install
    Start-Sleep -Seconds $SleepTimeInS
    write-host "Start Sleep Timer"
    
}
else
#______________________________________________________________________________________
#Else statemet for installing Windows updates through windows updates, and making external calls to other scripts to update 
{
    #______________________________________________________________________________________
    #Check Status of Windows Update service
    $WindowsUpdatateStatus = get-service -name wuauserv | Select-Object Name,Status,Starttype
    if($WindowsUpdatateStatus.Status -eq 'Stopped')
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
        Invoke-Expression $FlashUpdate
        write-host "Installed Flash Update"
    }
    
    #______________________________________________________________________________________
    #Check to see if Adobe Acrobat is installed
    if ((Get-Software | Select-Object DisplayName) -like "*Acrobat Reader*")
    {
        #______________________________________________________________________________________
        #Install Adobe Updates
        Invoke-Expression $AdobeUpdate
        write-host "Installed Adobe Acrobat Update"
    }
    
    #______________________________________________________________________________________
    #Check to see if Firefox is installed
    if ((Get-Software | Select-Object DisplayName) -like "*Mozilla Firefox*")
    {
        #______________________________________________________________________________________
        #Install Firefox Updates
        Invoke-Expression $FirefoxUpdate
        write-host "Installed Firefox Update"
    }
    
    #______________________________________________________________________________________
    #Check to see if Java is installed
    if ((Get-Software | Select-Object DisplayName) -like "*Java*")
    {
        #______________________________________________________________________________________
        #Install Java Updates
        Invoke-Expression $JavaUpdate
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
    Hide-WindowsUpdate -Title "Preview*" -Confirm:$false
    write-host "Hide Windows Updates labled as Preview"

    #______________________________________________________________________________________
    #Auto instal windows updates
    Get-WUInstall –AcceptAll -install
    write-host "Install Windows Updates"
    
    #______________________________________________________________________________________
    #Start Sleep timer for update installs.
    # Start-Sleep -Seconds $SleepTimeInS
    write-host "Wait for Windows Updates to install"

}

#______________________________________________________________________________________
#Decrypt Password for imput into reg.
$VDIMSTRTextCreds = $VDIMSTRCreds.Password | ConvertFrom-SecureString
$VDIMSTRTextCredsPlain = [Runtime.InteropServices.Marshal]::PtrToStringAuto([Runtime.InteropServices.Marshal]::SecureStringToBSTR( (ConvertTo-SecureString $VDIMSTRTextCreds) ))

#______________________________________________________________________________________
#Set AutoLog on and auto kick off step3
$RegPath = "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Winlogon"
$RegROPath = "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\RunOnce"
Set-ItemProperty $RegPath "AutoAdminLogon" -Value "1" -type String
write-host "Set Registry Key for AutoLogon"  
Set-ItemProperty $RegPath "DefaultUsername" -Value $VDIMSTRCreds.UserName -type String
write-host "Set Registry Key for AutoLogon Username"  
Set-ItemProperty $RegPath "DefaultPassword" -Value $VDIMSTRTextCredsPlain -type String
write-host "Set Registry Key for AutoLogon Password"
Set-ItemProperty $RegROPath "(Default)" -Value "c:\WINDOWS\system32\WindowsPowerShell\v1.0\powershell.exe -executionpolicy unrestricted -file $ShutdownScript" -type String
write-host "Set Registry Key for AutoLogon Logon Script"

#______________________________________________________________________________________
#Remote registry keys for legal notice
Remove-ItemProperty -Path "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\System" -Name "legalnoticecaption"
Write-Host "Remove Legal Notice Title Registry Key"
Remove-ItemProperty -Path "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\System" -Name "legalnoticetext"
Write-Host "Remove Legal Notice Text Registry Key "

#______________________________________________________________________________________
# Capture installed Software before updating
get-software | Format-Table | Out-File "$LogLocation\InstalledSoftwareAfter.txt"
write-host "Capture Installed Software After update and save to Logs Folder"

#______________________________________________________________________________________
# Check if Custom REG key is created if not create LastRecomposeDate REG key with 1/1/2000 date
if ( -not (Test-Path -Path $CustomRegPath)) {New-Item -path $CustomRegPath -Force | New-ItemProperty -Name  LastRecomposeDate -Value "01/01/2000" -Force | Out-Null}
write-host "Check if Custom key has been created and if not create registry key and set value to 1-1-2000"

#______________________________________________________________________________________
# Get LastRecomposeData value
$LastRecomposeDate = Get-ItemProperty -Path $CustomRegPath -Name LastRecomposeDate | Select-Object -ExpandProperty LastRecomposeDate
write-host "Get Last recompose date from custom Registry key, Last Recompose date was $LastRecomposeDate"

#______________________________________________________________________________________
# Get Windows Udates installed since last recompose. 
Get-HotFix | Where-Object -Property InstalledOn -GT $LastRecomposeDate | Out-File -FilePath "$LogLocation\WindowsUpdates.txt"
write-host "Create TXT file with Installed updates as of $LastRecomposeDate"

#______________________________________________________________________________________
#Set REG Key to store Last Recompose Date
Set-ItemProperty -Path $CustomRegPath -Name LastRecomposeDate -Value "$ScriptDate2" -Force
write-host "Set Last Recompose date to $ScriptDate2"

#______________________________________________________________________________________
# Stop Logging
Stop-Transcript

#______________________________________________________________________________________
#Reboot VM
Shutdown -r 