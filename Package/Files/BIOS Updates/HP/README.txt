The naming of the HP System BIOS files should be:
<Family>_<Version>.exe

<Model>
Elevated PowerShel prompt on the local machine:
(Get-WmiObject -Class Win32_ComputerSystem).Model
List of all models from an elevated PowerShell prompt on your SCCM server:
$Models = (Get-WmiObject -Namespace "root\SMS\SITE_$($SiteCode)" -Query "Select DISTINCT Version from SMS_G_System_COMPUTER_SYSTEM").Model

<Version>
The version is the System BIOS Version .

Naming Examples:
System BIOS Update file for HP ProDesk 600 G2
N02_02.37.bin