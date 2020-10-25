The naming of the Dell System BIOS files should be:
<Model>;<Model>;_<Version>.exe

<Model>
Elevated PowerShel prompt on the local machine:
(Get-WmiObject -Class Win32_ComputerSystem).Model
List of all models from an elevated PowerShell prompt on your SCCM server:
$Models = (Get-WmiObject -Namespace "root\SMS\SITE_$($SiteCode)" -Query "Select DISTINCT Version from SMS_G_System_COMPUTER_SYSTEM").Model

<Version>
The version is the System BIOS Version .

Naming Examples:
System BIOS file for Latitude E5270, Latitude E5470, Latitude E5570 and Precision 3510
Latitude E5270;Latitude E5470;Latitude E5570;Precision 3510;_1.19.3.exe

System BIOS file for Latitude E5440
Latitude E5440;_A23.exe