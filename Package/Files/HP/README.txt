The bios files are extracted using HPs bios exe.The naming of the HP System BIOS files should be:
<Model>;<Model>;_<Version>.bin

<Model>
Elevated PowerShel prompt on the local machine:
(Get-WmiObject -Class Win32_ComputerSystem).Model
List of all models from an elevated PowerShell prompt on your SCCM server:
$Models = (Get-WmiObject -Namespace "root\SMS\SITE_$($SiteCode)" -Query "Select DISTINCT Version from SMS_G_System_COMPUTER_SYSTEM").Model

<Version>
The version is the System BIOS Version .

Naming Examples:
System BIOS Update file for HP Z440 Workstation
 HP Z440 Workstation;_02.56.bin
