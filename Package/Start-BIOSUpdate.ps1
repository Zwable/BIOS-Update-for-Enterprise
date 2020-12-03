<#

    .SYNOPSIS
   
    .DESCRIPTION
    Use the service UI version that mathes the bitness in WinPE.
    Use the following commandline, in a commandline step in TS:
    ServiceUI_<bitness>.exe -process:winlogon.exe %SYSTEMROOT%\System32\WindowsPowerShell\v1.0\powershell.exe -NoProfile -WindowStyle Hidden -ExecutionPolicy Bypass -File .\<script>.ps1
    Use the following commandline in a package:
    %systemroot%\sysnative\wscript.exe .\<script>.vbs
    .PARAMETER

    .EXAMPLE

    .NOTES
    Author: Morten Rønborg
    Date: 28-03-2019
    Last Updated: 03-12-2020
    https://www.zwable.com

#>
################################################

param(
    $BiosFilesRoot = "$PSScriptRoot\Files"
)

#Dot Source
. "$PSScriptRoot\lib.ps1"

#Start transcript.
Start-Transcript -Path ("$env:SystemRoot\Logs\Software\BIOS-Update-" + (Get-Date).ToString("ddMMyyyy") + ".log") -Append -Force

#First check if the location excists
If(!(Test-Path $BiosFilesRoot)){
    Write-Host "No access to files folder '$($BiosFilesRoot), exiting.."
    #Stop transcript
    Stop-Transcript
    Exit 1337
}

#Temp folder for BIOS files
$TempFilesDestination = "$env:temp\BIOS Upgrade" #This folder will get deleted after the BIOS update has been performed!

#Get systeminfo
$SystemInfo = (Get-WmiObject -Namespace root\wmi MS_SystemInformation)

#Get the manufacturer of the running PC
$LocalManufacturer = $SystemInfo.SystemManufacturer

#Get the bios version of the running PC
[System.Version]$BIOSVersion = ([string]$SystemInfo.BiosMajorRelease + "." + [string]$SystemInfo.BiosMinorRelease)

#Declare hashtable
[hashtable]$Global:UpgradeInfo = @{}

#Add info to hashtable
if($LocalManufacturer -eq "Dell Inc."){

    #Define dell specific variables
    $Model = $SystemInfo.SystemProductName

    #Add to hashtable
    $UpgradeInfo.Add('ManufacturerFileRoot', "$BiosFilesRoot\Dell")

    if($BiosFile = (Get-ChildItem -Path "$($UpgradeInfo.ManufacturerFileRoot)\*$Model;*" -ErrorAction SilentlyContinue)){

        #Define info
        $UpgradeInfo.Add('Manufacturer', "Dell")
        $UpgradeInfo.Add('Model',  $Model)
        $UpgradeInfo.Add('CurrentBiosVersion', $BIOSVersion)
        $UpgradeInfo.Add('AvailableBiosVersion', [System.Version]([regex]::Matches(($($BiosFile.BaseName)).Split("_")[1], "\d+(\.\d+)+").Value))
        $UpgradeInfo.Add('BiosFilePath', $BiosFile.FullName)
        $UpgradeInfo.Add('TempRoot', $TempFilesDestination)
        $UpgradeInfo.Add('TempBiosFilePath', ("$TempFilesDestination\$($BiosFile.Name)"))
        $UpgradeInfo.Add('IgnoredExitCodes', $DellIgnoredExitcodes)
    }
#If manu is HP
}elseif($LocalManufacturer -eq "Hewlett-Packard"){

    #Define HP specific variables
    $Model = $SystemInfo.SystemProductName

    #Add to hashtable, HP uses a family alias for supporting multiple devices with the same file
    $UpgradeInfo.Add('ManufacturerFileRoot',  "$BiosFilesRoot\HP")

    if($BiosFile = (Get-ChildItem -Path "$($UpgradeInfo.ManufacturerFileRoot)\*$Model;*" -ErrorAction SilentlyContinue)){

        #Define info
        $UpgradeInfo.Add('Manufacturer', "HP")
        $UpgradeInfo.Add('Model', $Model)
        $UpgradeInfo.Add('CurrentBiosVersion', $BIOSVersion)
        $UpgradeInfo.Add('AvailableBiosVersion', [System.Version]([regex]::Matches(($($BiosFile.BaseName)).Split("_")[1], "\d+(\.\d+)+").Value))
        $UpgradeInfo.Add('BiosFilePath', $BiosFile.FullName)
        $UpgradeInfo.Add('HPBiosFlashToolPath', "$BiosFilesRoot\HP\HPBIOSUPDREC64.exe")
        $UpgradeInfo.Add('HPBiosFlashToolTempPath', "$TempFilesDestination\HPBIOSUPDREC64.exe")
        $UpgradeInfo.Add('TempRoot', $TempFilesDestination)
        $UpgradeInfo.Add('TempBiosFilePath', ("$TempFilesDestination\$($BiosFile.Name)"))
        $UpgradeInfo.Add('IgnoredExitCodes', $HPIgnoredExitcodes)
    }
#If manu is Lenovo
}elseif($LocalManufacturer -eq "Lenovo"){

    #Define Lenovo specific variables
    $Model = $SystemInfo.SystemVersion

    #Add to hashtable
    $UpgradeInfo.Add('ManufacturerFileRoot',  "$BiosFilesRoot\Lenovo")

    if($BiosFile = (Get-ChildItem -Path "$($UpgradeInfo.ManufacturerFileRoot)\*$Model;*" -ErrorAction SilentlyContinue)){

        #Define info
        $UpgradeInfo.Add('Manufacturer', "Lenovo")
        $UpgradeInfo.Add('Model', $Model)
        $UpgradeInfo.Add('CurrentBiosVersion', $BIOSVersion)
        $UpgradeInfo.Add('AvailableBiosVersion', [System.Version]([regex]::Matches(($($BiosFile.BaseName)).Split("_")[1], "\d+(\.\d+)+").Value))
        $UpgradeInfo.Add('BiosFilePath', $($BiosFile.FullName))
        $UpgradeInfo.Add('TempRoot', $TempFilesDestination)
        $UpgradeInfo.Add('TempBiosFilePath', ("$TempFilesDestination\$($BiosFile.Name)"))
        $UpgradeInfo.Add('IgnoredExitCodes', $LenovoIgnoredExitcodes)
    }
}

#Get last successful
[Globalization.CultureInfo]$Culture = Get-Culture
[nullable[datetime]]$LastBootUpTime = (Get-Date -ErrorAction 'Stop') - ([timespan]::FromMilliseconds([math]::Abs([Environment]::TickCount)))
$LastRunTimeRegKey = (Get-ItemProperty -Path "HKLM:\Software\BIOSUpdateForEnterprise" -Name 'Date' -ErrorAction 'SilentlyContinue').Date
if ($LastRunTimeRegKey) {
    [nullable[datetime]]$LastRunTime = [datetime]::Parse($LastRunTimeRegKey, $Culture) 

    #Checking if the machine has rebooted since the last successfull run
    if(!($LastBootUpTime -gt $LastRunTime)){

        #Exit the script without the users knowledge
        Write-Host "This machine was not rebooted (boot time $LastBootUpTime) since last successfull run (last run $LastRunTime), exiting"
        Stop-Transcript
        Exit 0
    }
}

#Write host
Write-Host "Running BIOS upgrade for the model '$Model'"

#If there is a file
if(($UpgradeInfo).BiosFilePath){
    
    Write-Host "Current installed version is '$($UpgradeInfo.CurrentBiosVersion)' and the available one is '$($UpgradeInfo.AvailableBiosVersion)'"

    #Brek if current installed version is greater or equals to the available one
    if($($UpgradeInfo.CurrentBiosVersion) -ge $($UpgradeInfo.AvailableBiosVersion)){

        #Write host
        Write-Host "Current installed version is compliant or higher than the available one, exiting..."
        #Stop transcript
        Stop-Transcript
        Exit 0
    }
    Write-Host "Current available version is higher than the one installed, proceeding..."
    
}else{
    
    #Write host
    Write-Host "No file found for this model, exiting.."
    #Stop transcript
    Stop-Transcript
    Exit 0
}

try {

    #First remove old folder
    Write-Host "Deleting potential old folder '$($UpgradeInfo.TempRoot)'"
    Remove-Item -Path $($UpgradeInfo.TempRoot) -Force -Recurse -Confirm:$false -ErrorAction SilentlyContinue | Out-Null
    
    #Copy the file locally in case the program is running from a fileshare or DP
    Write-Host "Creating content folder '$($UpgradeInfo.TempRoot)'"
    New-Item -ItemType Directory -Path $($UpgradeInfo.TempRoot) -Force | Out-Null

    #Copy the file locally in case the program is running from a fileshare or DP
    Write-Host "Copying file from '$($UpgradeInfo.BiosFilePath)' to '$($UpgradeInfo.TempRoot)'"
    Copy-Item -Path $($UpgradeInfo.BiosFilePath) -Destination $($UpgradeInfo.TempRoot) -Recurse | Out-Null

    if($UpgradeInfo.Manufacturer -eq "HP"){

        #If HP then include the flash tool
        Write-Host "Copying HP flash tool from '$($UpgradeInfo.HPBiosFlashToolPath)' to '$($UpgradeInfo.HPBiosFlashToolTempPath)'"
        Copy-Item -Path $($UpgradeInfo.HPBiosFlashToolPath) -Destination $($UpgradeInfo.HPBiosFlashToolTempPath) -Recurse | Out-Null
    }
}
catch {

    #Exit the script without the users knowledge
    Write-Host "Failed to create the local files: $_"
    Stop-Transcript
    Exit 1337
}

#Show message and start
Show-WPFMessage -Model $($UpgradeInfo.Model)
