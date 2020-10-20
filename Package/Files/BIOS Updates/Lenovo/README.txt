The naming of the Lenovo BIOS Update Utility files should be:
<Model>;<Model>;_<Version>.exe

<Model>
Elevated PowerShel prompt on the local machine:
(Get-WmiObject -Class Win32_ComputerSystemProduct).Version
List of all models from an elevated PowerShell prompt on your SCCM server:
$Models = (Get-WmiObject -Namespace "root\SMS\SITE_$($SiteCode)" -Query "Select DISTINCT Version from SMS_G_System_COMPUTER_SYSTEM_PRODUCT").Version

<Version>
The version is the BIOS package version .

Naming Examples:
BIOS Update Utility file for ThinkPad E480, ThinkPad E580 and ThinkPad R480
ThinkPad E480;ThinkPad E580;ThinkPad R480;_1.31.exe

BIOS Update Utility file for ThinkPad P50
ThinkPad P50;_1.56.exe