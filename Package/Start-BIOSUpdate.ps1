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
    Last Updated: 20-10-2019
    https://www.zwable.com

#>
################################################

param(
    $FilesRoot = "$PSScriptRoot\Files"
)

#Start transcript.
Start-Transcript -Path ("$env:SystemRoot\Logs\Software\BIOS-Update-" + (Get-Date).ToString("ddMMyyyy") + ".log") -Append -Force

#First check if the location excists
If(!(Test-Path $FilesRoot)){
    Write-Host "No access to files folder '$($FilesRoot), exiting.."
    #Stop transcript
    Stop-Transcript
    Exit 1337
}

# Dot Source library and define variables
. "$PSScriptRoot\lib.ps1"
$BiosFilesRoot = "$FilesRoot\BIOS Updates"
$TempFilesDestination = "$env:temp\BIOS Upgrade" #This folder will get deleted on exit!
$Global:MinimumBatteryPercentage = 80
$DellIgnoredExitcodes = @("2")
$HPIgnoredExitcodes = @("282")
$LenovoIgnoredExitcodes = @("-1","1")

#Get the manufacturer of the running PC
$LocalManufacturer = (Get-WmiObject -Class win32_bios).Manufacturer

#Get the bios version of the running PC
$SMBIOSBIOSVersion = (Get-WmiObject Win32_Bios).SMBIOSBIOSVersion

#Declare hashtable
[hashtable]$Global:UpgradeInfo = @{}

#Add info to hashtable
if($LocalManufacturer -eq "Dell Inc."){

    #Define dell specific variables
    $Model = (Get-WmiObject -Class Win32_ComputerSystem).Model

    #Add to hashtable
    $UpgradeInfo.Add('ManufacturerFileRoot', "$BiosFilesRoot\Dell")

    if($BiosFile = (Get-ChildItem -Path "$($UpgradeInfo.ManufacturerFileRoot)\*$Model;*" -ErrorAction SilentlyContinue)){

        #Dell has a mix of version numbers and reference numbers such as A08 or 1.3.12, can be handled by powershell comparer
        $UpgradeInfo.Add('Manufacturer', "Dell")
        $UpgradeInfo.Add('Model',  $Model)
        $UpgradeInfo.Add('CurrentBiosVersion', $SMBIOSBIOSVersion)
        $UpgradeInfo.Add('AvailableBiosVersion', (($BiosFile).BaseName).Split("_")[1])
        $UpgradeInfo.Add('BiosFilePath', $BiosFile.FullName)
        $UpgradeInfo.Add('TempRoot', $TempFilesDestination)
        $UpgradeInfo.Add('TempBiosFilePath', ("$TempFilesDestination\$($BiosFile.Name)"))
        $UpgradeInfo.Add('IgnoredExitCodes', $DellIgnoredExitcodes)
    }
#If manu is HP
}elseif($LocalManufacturer -eq "Hewlett-Packard"){

    #Define HP specific variables
    $Model = (Get-WmiObject -Class Win32_ComputerSystem).Model

    #Add to hashtable, HP uses a family alias for supporting multiple devices with the same file
    $UpgradeInfo.Add('ManufacturerFileRoot',  "$BiosFilesRoot\HP")
    $BIOSFamily = ($SMBIOSBIOSVersion).Split(" ")[0]

    if($BiosFile = (Get-ChildItem -Path "$($UpgradeInfo.ManufacturerFileRoot)\*$BIOSFamily*" -ErrorAction SilentlyContinue)){

        #Define info
        $UpgradeInfo.Add('Manufacturer', "HP")
        $UpgradeInfo.Add('Model', $Model)
        $UpgradeInfo.Add('CurrentBiosVersion', [System.Version]([regex]::Matches($SMBIOSBIOSVersion, "\d+(\.\d+)+").Value))
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
    $Model = (Get-WmiObject -Class Win32_ComputerSystemProduct).Version

    #Add to hashtable
    $UpgradeInfo.Add('ManufacturerFileRoot',  "$BiosFilesRoot\Lenovo")

    if($BiosFile = (Get-ChildItem -Path "$($UpgradeInfo.ManufacturerFileRoot)\*$Model;*" -ErrorAction SilentlyContinue)){

        #Define info
        $UpgradeInfo.Add('Manufacturer', "Lenovo")
        $UpgradeInfo.Add('Model', $Model)
        $UpgradeInfo.Add('CurrentBiosVersion', ([regex]::Matches($SMBIOSBIOSVersion, "\d+(\.\d+)+").Value) )
        $UpgradeInfo.Add('AvailableBiosVersion', ([regex]::Matches(($($BiosFile.BaseName)).Split("_")[1], "\d+(\.\d+)+").Value))
        $UpgradeInfo.Add('BiosFilePath', $($BiosFile.FullName))
        $UpgradeInfo.Add('TempRoot', $TempFilesDestination)
        $UpgradeInfo.Add('TempBiosFilePath', ("$TempFilesDestination\$($BiosFile.Name)"))
        $UpgradeInfo.Add('IgnoredExitCodes', $LenovoIgnoredExitcodes)
    }
}

#Write host
Write-Host "Setting BIOS settings for the model '$Model'"

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

#Define text
$HeaderText = $($UpgradeInfo.Model)
$BodyText = 
@"
Due to security and system stability the BIOS/firmware of your PC needs to be patched.

As BIOS is stored in the flash memory of the baseboard chip, the installation will continue on the next reboot (which is not enforced), after this installation.
"@
$BottomText = "Installation will automatically start when requirements are met:`n-More than $MinimumBatteryPercentage`% of battery charged.`n-AC adapter plugged in."

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
Show-WPFMessage