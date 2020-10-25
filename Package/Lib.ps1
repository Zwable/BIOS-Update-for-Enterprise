<#

    .SYNOPSIS
    Library file

    .DESCRIPTION

    .PARAMETER

    .EXAMPLE

    .NOTES
    Author: Morten Rønborg
    Date: 04-12-2018
    Last Updated: 21-10-2020

#>

################################################

#Import variables from XML configuration file
$CurrentLoggedOnUserSID = (New-Object -ComObject Microsoft.DiskQuota).TranslateLogonNameToSID((Get-WmiObject -Class Win32_ComputerSystem).Username)
New-PSDrive HKU -PSProvider Registry -Root HKEY_USERS
$UICultureShortname = Get-ItemPropertyValue "HKU:$CurrentLoggedOnUserSID\Control Panel\Desktop" "PreferredUILanguages"
Remove-PSDrive HKU
[Xml.XmlDocument]$xmlConfigFile = Get-Content -LiteralPath "$PSScriptRoot\BiosUpdateConfig.xml"
[Xml.XmlElement]$xmlConfig = $xmlConfigFile.BiosUpdate_Config
$xmlUIMessageLanguage = "UI_Messages_$UICultureShortname"
If (!($xmlConfig.$xmlUIMessageLanguage)) {$xmlUIMessageLanguage = 'UI_Messages_en-US'}

#Define option variables
$Global:MinimumBatteryPercentage = $xmlConfig.BiosUpdateConfig_Options.MinimumBatteryPercentage
[string[]]$Global:DellIgnoredExitcodes = ($xmlConfig.BiosUpdateConfig_Options.DellIgnoredExitcodes).Split(",")
[string[]]$Global:HPIgnoredExitcodes = ($xmlConfig.BiosUpdateConfig_Options.HPIgnoredExitcodes).Split(",")
[string[]]$Global:LenovoIgnoredExitcodes = ($xmlConfig.BiosUpdateConfig_Options.LenovoIgnoredExitcodes).Split(",")

#Define UI variables
[hashtable]$Global:GUIText = @{}
$GUIText.Add('TitleText', $xmlConfig.$xmlUIMessageLanguage.TitleText)
$GUIText.Add('HeaderText', $xmlConfig.$xmlUIMessageLanguage.HeaderText)
$GUIText.Add('BodyText', $xmlConfig.$xmlUIMessageLanguage.BodyText)
$GUIText.Add('BottomText', $xmlConfig.$xmlUIMessageLanguage.BottomText)
$GUIText.Add('NotChargingText', $xmlConfig.$xmlUIMessageLanguage.NotChargingText)
$GUIText.Add('ChargingText', $xmlConfig.$xmlUIMessageLanguage.ChargingText)
$GUIText.Add('InstallingText', $xmlConfig.$xmlUIMessageLanguage.InstallingText)
$GUIText.Add('BalloonTipTitle', $xmlConfig.$xmlUIMessageLanguage.BalloonTipTitle)
$GUIText.Add('BalloonTipText', $xmlConfig.$xmlUIMessageLanguage.BalloonTipText)
$GUIText.Add('InstallSuccessText', $xmlConfig.$xmlUIMessageLanguage.InstallSuccessText)
$GUIText.Add('ErrorText', $xmlConfig.$xmlUIMessageLanguage.ErrorText)

function Start-LoadAssemblys{

    param(
        $RootPath
    )

    #Add assemblys
    Add-Type -AssemblyName PresentationFramework, System.Drawing, System.Windows.Forms, WindowsFormsIntegration
    #Add assembly using unsafe method. To support network location in case running from DP/Share read as bytes to not lock files
    $DLLBytes = [System.IO.File]::ReadAllBytes("$RootPath\MaterialDesignThemes.Wpf.dll")
    [System.Reflection.Assembly]::Load($DLLBytes)
    $DLLBytes = [System.IO.File]::ReadAllBytes("$RootPath\MaterialDesignColors.dll")
    [System.Reflection.Assembly]::Load($DLLBytes)
    $DLLBytes = [System.IO.File]::ReadAllBytes("$RootPath\PresentationFramework.dll")
    [System.Reflection.Assembly]::Load($DLLBytes)
}


Function Get-MainWindowXAML{

    #Define the XAML
    [xml]$XAML = @"
    <Window x:Class="MainWindow"
    TextElement.Foreground="{DynamicResource MaterialDesignBody}"
    xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
    xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
    xmlns:d="http://schemas.microsoft.com/expression/blend/2008"
    xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
    xmlns:materialDesign="http://materialdesigninxaml.net/winfx/xaml/themes"
    mc:Ignorable="d"
    FontSize="15"
    Title="FirmwareUpgrade" Height="450" Width="500" WindowStartupLocation="Manual" ResizeMode="NoResize" ShowInTaskbar="False" Topmost="True" AllowsTransparency="True" WindowStyle="None">
    <Window.Resources>
        <ResourceDictionary>
            <ResourceDictionary.MergedDictionaries>
                <ResourceDictionary Source="pack://application:,,,/MaterialDesignThemes.Wpf;component/Themes/MaterialDesignTheme.Light.xaml" />
                <ResourceDictionary Source="pack://application:,,,/MaterialDesignThemes.Wpf;component/Themes/MaterialDesignTheme.Defaults.xaml" />
                <ResourceDictionary Source="pack://application:,,,/MaterialDesignColors;component/Themes/Recommended/Primary/MaterialDesignColor.Blue.xaml" />
                <ResourceDictionary Source="pack://application:,,,/MaterialDesignColors;component/Themes/Recommended/Accent/MaterialDesignColor.Blue.xaml" />
            </ResourceDictionary.MergedDictionaries>
        </ResourceDictionary>
    </Window.Resources>
    <Grid x:Name="grd_mainwindow">
        <Grid.RowDefinitions>
            <RowDefinition Height="60"/>
            <RowDefinition Height="*"/>
            <RowDefinition Height="90"/>
            <RowDefinition Height="100"/>
        </Grid.RowDefinitions>
        <Grid Grid.Row ="0" Background="#106FAA">
            <StackPanel Orientation="Horizontal" Grid.ColumnSpan="3">
                <materialDesign:PackIcon Kind="Computer" Margin="0,10" VerticalAlignment="Stretch" Height="Auto" Width="60"/>
                <TextBlock x:Name="tbx_title" Text="BIOS/Firmware upgrade" FontWeight="Bold" FontSize="25" HorizontalAlignment="Left" VerticalAlignment="Center"/>
            </StackPanel>
            <Button x:Name="btn_minimize" Content="_" FontSize="12" VerticalAlignment="Top" HorizontalAlignment="Right"  Background="#106FAA" />
            <!--<Image HorizontalAlignment="right" x:Name="img_logo" Width="122" Margin="0,10,10,10"/>-->
        </Grid>
        <Grid Grid.Row ="1">
            <Grid.RowDefinitions>
                <RowDefinition Height="30"/>
                <RowDefinition Height="*"/>
            </Grid.RowDefinitions>
            <Grid Grid.Row="0">
                <Label FontWeight="Bold" FontSize="18" Background="LightGray">
                    <TextBlock x:Name="tbx_header" Text="ScrollViewHeader" TextWrapping="Wrap"/>
                </Label>
            </Grid>
            <Grid Grid.Row="1" Background="LightGray">
                <ScrollViewer VerticalScrollBarVisibility="Auto">
                    <Label FontWeight="DemiBold" FontSize="17">
                        <TextBlock x:Name="tbx_body" Text="ScrollViewBody" VerticalAlignment="Top" TextWrapping="Wrap"/>
                    </Label>
                </ScrollViewer>
            </Grid>
        </Grid>
        <Grid Grid.Row="2" Background="LightGray">
            <ScrollViewer VerticalScrollBarVisibility="Auto">
                <Label FontWeight="DemiBold" FontSize="15">
                    <TextBlock x:Name="tbx_message" Text="ScrollViewMessage" VerticalAlignment="Top" TextWrapping="Wrap" FontWeight="Bold"/>
                </Label>
            </ScrollViewer>
        </Grid>
        <Grid Grid.Row ="3" x:Name="grd_batteryinfo" Visibility="Visible" Background="LightGray">
            <materialDesign:Card VerticalAlignment="Stretch" HorizontalAlignment="Stretch" Margin="10 10 10 10" materialDesign:ShadowAssist.ShadowDepth="Depth2">
                <StackPanel Orientation="Horizontal">
                    <materialDesign:PackIcon x:Name="ico_status" Kind="BatteryOutline" VerticalAlignment="Stretch" HorizontalAlignment="Stretch" Margin="0,10" Height="Auto" Width="75">
                    </materialDesign:PackIcon>
                    <TextBlock x:Name="tbx_percentage" Text="??%" FontSize="17" VerticalAlignment="Center" HorizontalAlignment="Center" FontWeight="Bold"/>
                    <TextBlock x:Name="tbx_percentagetext" Text="" FontSize="17" VerticalAlignment="Center" HorizontalAlignment="Center" FontWeight="Bold" Margin="10 10 10 10"/>
                </StackPanel>
            </materialDesign:Card>
        </Grid>
        <Grid Grid.Row ="3" x:Name="grd_installationdone" Visibility="Hidden" Background="LightGray">
            <materialDesign:Card VerticalAlignment="Stretch" HorizontalAlignment="Stretch" Margin="10 10 10 10" materialDesign:ShadowAssist.ShadowDepth="Depth2">
                <StackPanel Orientation="Horizontal" Width="auto">
                    <materialDesign:PackIcon x:Name="ico_installationdone" Kind="Check" VerticalAlignment="Stretch" HorizontalAlignment="Stretch" Margin="0,10" Height="Auto" Width="75">
                    </materialDesign:PackIcon>
                    <TextBlock Width="auto" TextWrapping="Wrap" x:Name="tbx_installationdone" Text="Reboot to complete the installation" VerticalAlignment="Center" HorizontalAlignment="Center" FontWeight="Bold" Margin="10 10 10 10"/>
                    <Button x:Name="btn_exit" Content="Exit" FontSize="12" VerticalAlignment="Center" HorizontalAlignment="Center" Background="#106FAA" Margin="50 0 0 0"/>
                </StackPanel>
            </materialDesign:Card>
        </Grid>
        <materialDesign:DialogHost x:Name="pop_installfirmware" IsOpen="False" Grid.RowSpan="5">
            <materialDesign:DialogHost.DialogContent>
                <StackPanel Margin="29" Orientation="Vertical" Width="120" Height="160">
                    <TextBlock x:Name="tbx_installing" Text="Do not turn off your PC. Updating BIOS/firmware.." FontSize="15" FontWeight="Bold" TextWrapping="Wrap" TextAlignment="Center" Margin="0 0 0 0"/>
                    <StackPanel Orientation="Vertical">
                        <ProgressBar Style="{StaticResource MaterialDesignCircularProgressBar}"
                            Width="100"
                            IsIndeterminate="True"/>
                    </StackPanel>
                </StackPanel>
            </materialDesign:DialogHost.DialogContent>
        </materialDesign:DialogHost>
    </Grid>
</Window>
"@

#Remove attributes
$CleanedXAML = Get-CleanXAML -XAML $XAML

#Return
Return $CleanedXAML

}

Function Get-CleanXAML{
    param(
        $XAML
    )

    #Attributes to remove
    $AttributesToRemove = @(
        'x:Class',
        'mc:Ignorable'
    )

    #Remove XML attributes that breaks PS
    foreach ($Attrib in $AttributesToRemove) {
        if ( $XAML.Window.GetAttribute($Attrib) ) {
                $XAML.Window.RemoveAttribute($Attrib)
        }
    }

    #Return cleaned XML
    Return $XAML
}

function Get-IsLaptop{

    #Get type
    $Type =  (Get-WmiObject -Class Win32_ComputerSystem).PCSystemType

    if($Type -eq 2){
        $Laptop = $true
    }else{
        $Laptop = $false
    }

    #Return the build
    return $Laptop
}
function Start-CheckBattery{

    #Timer
    $Global:CheckBatteryTimer = New-Object System.Windows.Forms.Timer
    $CheckBatteryTimer.Interval=2000
    $CheckBatteryTimer.add_Tick({

        #Variables
        $MinimumPercentage = $Global:MinimumBatteryPercentage
        $BatteryInfo = Get-BatteryInfo
        $CurrentCharge = [math]::Round(($BatteryInfo).BatteryLifePercent * 100) #add 1 to get current charge
        $ACConnected = ($BatteryInfo).IsUsingACPower
        $BatteryIndicator = if($CurrentCharge -lt 10){"Outline"}else{($CurrentCharge -replace ".$","0")} #replace the last number with 0

        #Ivoke the action
        #If battery connected
        if($ACConnected){
            $WPFWindow.Window.Dispatcher.invoke([action]{
                $WPFWindow.tbx_percentage.Text ="$CurrentCharge`%"
                $WPFWindow.tbx_percentagetext.Text = if($CurrentCharge -lt 100){$GUIText.ChargingText}else{""}
                $WPFWindow.ico_status.Kind ="BatteryCharging$BatteryIndicator"
            },"Normal")
        }else{
            $WPFWindow.Window.Dispatcher.invoke([action]{
                $WPFWindow.tbx_percentage.Text = "$CurrentCharge`%"
                $WPFWindow.tbx_percentagetext.Text = $GUIText.NotChargingText
                $WPFWindow.ico_status.Kind ="Battery$BatteryIndicator"
            },"Normal")
        }

        #Check if the battery is charged and the ac is connected
        if((([int]$CurrentCharge -gt [int]$MinimumPercentage) -and $ACConnected) -or !(Get-IsLaptop)){

            #Run the installation
            if(!( Get-Job -State Running | Where-Object {$_.Name -eq "UpgradeBIOS"}) -and !( Get-Job -State Completed | Where-Object {$_.Name -eq "UpgradeBIOS"})){
                
                #Define the scriptblock
                $ScriptBlock={
                    Param(
                        $UpgradeInfo
                    )

                    #Variables
                    $ExtractLocation = "$($UpgradeInfo.TempRoot)\BIOSExtracted"
                    [environment]::CurrentDirectory = "C:\Windows\Temp"

                    if($($UpgradeInfo.Manufacturer) -eq "Dell"){

                        #Run the upgrade
                        $Process = Start-Process -FilePath $($UpgradeInfo.TempBiosFilePath) -ArgumentList "/s /f" -WorkingDirectory $($UpgradeInfo.TempRoot) -WindowStyle Hidden -PassThru -Wait
                        If($Process.ExitCode -ne 0){throw $Process.ExitCode}

                    }elseif($($UpgradeInfo.Manufacturer) -eq "HP"){
                      
                        #Run the upgrade
                        Write-Host "Executing the BIOS upgrade with '$($UpgradeInfo.HPBiosFlashToolTempPath)'"
                        $Process = Start-Process -FilePath $($UpgradeInfo.HPBiosFlashToolTempPath) -ArgumentList "-s -r" -WorkingDirectory $($UpgradeInfo.TempRoot) -WindowStyle Hidden -PassThru -Wait
                        If($Process.ExitCode -ne 0){throw $Process.ExitCode}

                    }elseif($($UpgradeInfo.Manufacturer) -eq "Lenovo"){

                        #Create the folder for extracted content
                        Write-Host "Creating extraction folder '$ExtractLocation'"
                        New-Item -ItemType Directory -Path $ExtractLocation -Force | Out-Null

                        #Extract package content
                        Write-Host "Extracting the file '$($UpgradeInfo.TempBiosFilePath)' to '$ExtractLocation'"
                        $Process = Start-Process -FilePath $($UpgradeInfo.TempBiosFilePath) -ArgumentList "/VERYSILENT /DIR=`"$ExtractLocation`" /Extract=`"YES`"" -WorkingDirectory $($UpgradeInfo.TempRoot) -WindowStyle Hidden -PassThru -Wait
                        If($Process.ExitCode -ne 0){throw $Process.ExitCode}
                        
                        #Run the upgrade
                        Write-Host "Executing the BIOS upgrade with '$ExtractLocation\WINUPTP.exe'"
                        $Process = Start-Process -FilePath "$ExtractLocation\WINUPTP.exe" -ArgumentList "-s" -WorkingDirectory $ExtractLocation -WindowStyle Hidden -PassThru -Wait
                        If($Process.ExitCode -ne 0){throw $Process.ExitCode}
                    }
                }

                #Start the job
                if(!($HasRun)){

                    $WPFWindow.Window.Dispatcher.Invoke([action]{
                        #Write host
                        Write-Host "Requirements reached, installing BIOS..."

                        #Invoke loading screen
                        $WPFWindow.Window.Topmost = $false
                        $WPFWindow.Window.IsEnabled = $false
                        $WPFWindow.Window.Visibility = "Visible"
                        $WPFWindow.tbx_installing.Text = $GUIText.InstallingText
                        $WPFWindow.pop_installfirmware.IsOpen = $true

                        #Suspend bitlocker if enabled
                        Set-BitlockerSuspend

                        #Start async job
                        Start-Job -ScriptBlock $ScriptBlock -Name UpgradeBIOS -ArgumentList @($UpgradeInfo)
                        
                    },"Normal")
                }

                #Ensure that it will not be invoked again
                $Global:HasRun = $true

            }

            if(Get-Job -State "Completed"){
                Get-Job | Receive-Job -Keep -ErrorAction SilentlyContinue -OutVariable ReturnSuccess
                [int]$Global:ReturnCode = "$($ReturnSuccess[0])"
                Write-Host "Succeded with :$($ReturnCode)"
                Start-Cleanup
                Set-LastSuccessfulRunDate
                Stop-Timers
                $WPFWindow.tbx_installationdone.Text = $GUIText.InstallSuccessText
                $WPFWindow.pop_installfirmware.IsOpen = $false
                $WPFWindow.Window.IsEnabled = $true
                $WPFWindow.Window.Topmost = $true
                $WPFWindow.grd_installationdone.Visibility = "Visible"
                $WPFWindow.grd_batteryinfo.Visibility = "Hidden"
                $WPFWindow.tbx_installationdone.Text = $GUIText.InstallSuccessText


            }elseif(Get-Job -State "Failed"){
                Get-Job | Receive-Job -Keep -ErrorAction SilentlyContinue -ErrorVariable ReturnErrors
                [int]$Global:ReturnCode = "$($ReturnErrors[0])"
                Write-Host "Failed with: $($ReturnCode)"
                Start-Cleanup
                Stop-Timers
                $WPFWindow.pop_installfirmware.IsOpen = $false
                $WPFWindow.Window.IsEnabled = $true
                $WPFWindow.Window.Topmost = $true

                #Change grid with faild text and not in ignored exitcodes
                if($($ReturnCode) -notin $($UpgradeInfo.IgnoredExitCodes)){
                    $WPFWindow.ico_installationdone.Kind = "Error"
                    $WPFWindow.tbx_installationdone.Text = ($GUIText.ErrorText -f $($ReturnCode))
                }else{
                    $WPFWindow.tbx_installationdone.Text = $GUIText.InstallSuccessText
                    Set-LastSuccessfulRunDate
                }

                $WPFWindow.grd_installationdone.Visibility = "Visible"
                $WPFWindow.grd_batteryinfo.Visibility = "Hidden"
            }

            #If running in task sequence and returncode is defined, exit without prompt
            if($RunningTaskSequence -and $ReturnCode){
                Clear-AndClose
            }
        }

    })

    #Start the timer
    $CheckBatteryTimer.Start()
}
Function Get-BatteryInfo {

        #Logic is taken from PSADT
		## Initialize a hashtable to store information about system type and power status
        [hashtable]$SystemTypePowerStatus = @{ }
        
        [Windows.Forms.PowerStatus]$PowerStatus = [Windows.Forms.SystemInformation]::PowerStatus
        ## Get the system power status. Indicates whether the system is using AC power or if the status is unknown. Possible values:
        [string]$PowerLineStatus = $PowerStatus.PowerLineStatus
		$SystemTypePowerStatus.Add('ACPowerLineStatus', $PowerStatus.PowerLineStatus)
		
		## Get the current battery charge status. Possible values: High, Low, Critical, Charging, NoSystemBattery, Unknown.
		[string]$BatteryChargeStatus = $PowerStatus.BatteryChargeStatus
		$SystemTypePowerStatus.Add('BatteryChargeStatus', $PowerStatus.BatteryChargeStatus)
		
		## Get the approximate amount, from 0.00 to 1.0, of full battery charge remaining.
		[single]$BatteryLifePercent = $PowerStatus.BatteryLifePercent
		If (($BatteryChargeStatus -eq 'NoSystemBattery') -or ($BatteryChargeStatus -eq 'Unknown')) {
			[single]$BatteryLifePercent = 0.0
		}
		$SystemTypePowerStatus.Add('BatteryLifePercent', $PowerStatus.BatteryLifePercent)
		
		## The reported approximate number of seconds of battery life remaining. It will report –1 if the remaining life is unknown because the system is on AC power.
        $SystemTypePowerStatus.Add('BatteryLifeRemaining', $PowerStatus.BatteryLifeRemaining)
        
		## Get the manufacturer reported full charge lifetime of the primary battery power source in seconds.
		$SystemTypePowerStatus.Add('BatteryFullLifetime', $PowerStatus.BatteryFullLifetime)
		
		## Determine if the system is using AC power
		[boolean]$OnACPower = $false
		If (($PowerLineStatus -eq 'Online') -or ($PowerLineStatus -eq 'Unknown')) {
			$OnACPower = $true
        }
        $SystemTypePowerStatus.Add('IsUsingACPower', $OnACPower)

        #Return the info
        Return $SystemTypePowerStatus
        
}
function Stop-Timers{

    #Stop timers
    if($CheckBatteryTimer.Enabled -eq $true){
        Write-Host "Stopping timers"
        $CheckBatteryTimer.Stop()
        $CheckBatteryTimer.Dispose()
    }
}

function Set-LastSuccessfulRunDate{

    #Setting key to prevent re-run before reboot (setting key in UniversalSortableDateTimePattern for multi culture support)
    [Globalization.CultureInfo]$Culture = Get-Culture
    [datetime]$DateTimeNow = $([DateTime]::Now)
    New-Item "HKLM:\Software\BIOSUpdateForEnterprise" -Force | New-ItemProperty -Name Date -Value ((Get-Date -Date $DateTimeNow -Format (($Culture).DateTimeFormat.UniversalSortableDateTimePattern)) -replace 'Z$', '') -Force | Out-Null
}

function Start-Cleanup{

    #Remove the BIOS files from the temp folder
    Write-Host "Deleting temporary folder '$TempFilesDestination'..."
    Remove-Item -Path $TempFilesDestination -Force -Recurse -Confirm:$false | Out-Null
}
function Clear-AndClose{
    
    #Stop timers
    Stop-Timers

    #Cleanup
    Start-Cleanup

    #Unreg eventhandler
    if($RunningTaskSequence -eq $false){

        Write-Host "Unregistering icon clicked eventhandler and removing job..."
        #Unregister event
        Unregister-Event  -SourceIdentifier IconClicked

        #Remove job
        Remove-Job -Name IconClicked
    }

    #Close window
    Write-Host "Closing form..."
    $WPFWindow.Window.Close()

    #Stop transcript
    Stop-Transcript
    
    #Exit with code, we use 2 different types to handle TS
    Write-Host "Returning exitcode '$ReturnCode'..."
    if($RunningTaskSequence -eq $false){
        Exit $ReturnCode
    }else{
        [System.Environment]::Exit($ReturnCode)
    }
}

function Set-BitlockerSuspend{

    # Detect Bitlocker Status
    $OSDriveEncrypted = $false
    $EncryptedVolumes = Get-WmiObject -Namespace "root\cimv2\Security\MicrosoftVolumeEncryption" -Class "Win32_EncryptableVolume"
    foreach ($Volume in $EncryptedVolumes) {
        if ($Volume.DriveLetter -like $env:SystemDrive) {
            if ($Volume.EncryptionMethod -ge 1) {
                $OSDriveEncrypted = $true
            }
        }
    }
            
    # Supend Bitlocker if $OSVolumeEncypted is $true
    if ($OSDriveEncrypted -eq $true) {
        Write-Host "Suspending BitLocker protected volume: $($env:SystemDrive)"
        Manage-Bde -Protectors -Disable $($env:SystemDrive)
    }	
}

function Set-WindowWorkArea{

    $ScreenWorkArea = [System.Windows.SystemParameters]::WorkArea
    $WPFWindow.Window.Top = ($ScreenWorkArea.Height - $WPFWindow.Window.Height)
    $WPFWindow.Window.Left = ($ScreenWorkArea.Width - $WPFWindow.Window.Width)
}
Function Show-WPFMessage{

    param(
        $Model
    )

    #Load assemblys
    Start-LoadAssemblys -RootPath $PSScriptRoot

    #Check if script is running from a SCCM Task Sequence
    Try {
        $SMSTSEnvironment = New-Object -ComObject 'Microsoft.SMS.TSEnvironment' -ErrorAction SilentlyContinue | Out-Null
        Write-Host "Successfully loaded COM Object [Microsoft.SMS.TSEnvironment]. Therefore, script is currently running from a SCCM Task Sequence."
        $Global:RunningTaskSequence = $true
    }
    Catch {
        Write-Host "Unable to load COM Object [Microsoft.SMS.TSEnvironment]. Therefore, script is not currently running from a SCCM Task Sequence."
        $Global:RunningTaskSequence = $false
    }
    
    #Define the XAML
    $MainWindowXAML = Get-MainWindowXAML

    #Read the XAML
    $MainWindowReader = (New-Object System.Xml.XmlNodeReader $MainWindowXAML)

    #Define the shared runspace hastable
    $Global:WPFWindow = @{}
    $WPFWindow.Window = [Windows.Markup.XamlReader]::Load($MainWindowReader)

    #Add to the hastable
    $MainWindowXAML.SelectNodes("//*[@*[contains(translate(name(.),'n','N'),'Name')]]") | ForEach-Object{
        
        #Find all of the form types and add them as members to the WPFWindow
        $WPFWindow.Add($_.Name,$WPFWindow.Window.FindName($_.Name) )
    }

    #Enble functions if not running from task sequnce
    if($RunningTaskSequence -eq $false){

        #Icontray and baloontip
        $NotifyIcon = New-Object System.Windows.Forms.NotifyIcon
        $NotifyIcon.Icon = [System.Convert]::FromBase64String("AAABAAQAICAAAAEAIACoEAAARgAAACAgAAABAAgAqAgAAO4QAAAQEAAAAQAgAGgEAACWGQAAEBAAAAEACABoBQAA/h0AACgAAAAgAAAAQAAAAAEAIAAAAAAAACAAAAAAAAAAAAAAAAAAAAAAAACqbxD/qm8Q/6pvEP+qbxD/qm8Q/6pvEP+qbxD/qm8Q/6pvEP+qbxD/qm8Q/6pvEP+qbxD/qm8Q/6pvEP+qbxD/qm8Q/6pvEP+qbxD/qm8Q/6pvEP+qbxD/qm8Q/6pvEP+qbxD/qm8Q/6pvEP+qbxD/qm8Q/6pvEP+qbxD/qm8Q/6pvEP+qbxD/qm8Q/6pvEP+qbxD/qm8Q/6pvEP+qbxD/qm8Q/6pvEP+qbxD/qm8Q/6pvEP+qbxD/qm8Q/6pvEP+qbxD/qm8Q/6pvEP+qbxD/qm8Q/6pvEP+qbxD/qm8Q/6pvEP+qbxD/qm8Q/6pvEP+qbxD/qm8Q/6pvEP+qbxD/qm8Q/6pvEP+qbxD/qm8Q/6pvEP+qbxD/qm8Q/6pvEP+qbxD/qm8Q/6pvEP+qbxD/qm8Q/6pvEP+qbxD/qm8Q/6pvEP+qbxD/qm8Q/6pvEP+qbxD/qm8Q/6pvEP+qbxD/qm8Q/6pvEP+qbxD/qm8Q/6pvEP+qbxD/qm8Q/6pvEP+qbxD/qm8Q/6pvEP+qbxD/qm8Q/6pvEP+qbxD/qm8Q/6pvEP+qbxD/qm8Q/6pvEP+qbxD/qm8Q/6pvEP+qbxD/qm8Q/6pvEP+qbxD/qm8Q/6pvEP+qbxD/qm8Q/6pvEP+qbxD/qm8Q/6pvEP+qbxD/qm8Q/6pvEP+qbxD/qm8Q/6pvEP+qbxD/qm8Q/6pvEP+qbxD/qm8Q/6pvEP+qbxD/qm8Q/6pvEP+qbxD/qm8Q/6pvEP+qbxD/qm8Q/6pvEP+qbxD/qm8Q/6pvEP+qbxD/qm8Q/6pvEP+qbxD/qm8Q/6pvEP+qbxD/qm8Q/6pvEP+qbxD/qm8Q/6pvEP+qbxD/qm8Q/6pvEP+qbxD/qm8Q/6pvEP+qbxD/qm8Q/6pvEP+qbxD/qm8Q/6pvEP+qbxD/qm8Q/6pvEP+qbxD/qm8Q/6pvEP+qbxD/qm8Q/6pvEP+qbxD/qm8Q/6pvEP+qbxD/qm8Q/6pvEP+qbxD/qm8Q/6pvEP+qbxD/qm8Q/6pvEP+qbxD/qm8Q/6pvEP+qbxD/qm8Q/6pvEP+qbxD/qm8Q/6pvEP+qbxD/qm8Q/6pvEP+qbxD/qm8Q/6pvEP+qbxD/qm8Q/6pvEP+qbxD/qm8Q/6pvEP+qbxD/qm8Q/6pvEP+qbxD/qm8Q/6pvEP+qbxD/qm8Q/6pvEP+qbxD/qm8Q/6pvEP+qbxD/qm8Q/6pvEP+qbxD/qm8Q/6pvEP+qbxD/qm8Q/6pvEP+qbxD/qm8Q/6pvEP+qbxD/qm8Q/6pvEP+qbxD/qm8Q/6pvEP+qbxD/qm8Q/6pvEP+qbxD/qm8Q/6pvEP+qbxD/qm8Q/6pvEP+qbxD/qm8Q/6pvEP+qbxD/qm8Q/6pvEP+qbxD/qm8Q/6pvEP98UAv/XTwI/108CP9dPAj/XTwI/108CP9dPAj/XTwI/108CP9dPAj/XTwI/108CP9dPAj/XTwI/108CP9dPAj/XTwI/108CP9dPAj/XTwI/108CP9dPAj/oGgO/6pvEP+qbxD/qm8Q/6pvEP+qbxD/qm8Q/6pvEP+qbxD/qm8Q/1M2B/8XDwL/Fw8C/xcPAv8XDwL/Fw8C/xcPAv8XDwL/Fw8C/xcPAv8XDwL/Fw8C/xcPAv8XDwL/Fw8C/xcPAv8XDwL/Fw8C/xcPAv8XDwL/Fw8C/xcPAv+XYg7/qm8Q/6pvEP+qbxD/qm8Q/6pvEP+qbxD/qm8Q/6pvEP+qbxD/j10N/31SC/9sRgn/KRsD/xcPAv8XDwL/Fw8C/xcPAv8XDwL/Fw8C/xcPAv8XDwL/Fw8C/xcPAv8XDwL/Fw8C/xcPAv8XDwL/GBAC/0QsBv96Twv/fVIL/6RrD/+qbxD/qm8Q/6pvEP+qbxD/qm8Q/6pvEP+qbxD/qm8Q/6pvEP+qbxD/qm8Q/1E0B/8YDwL/HhMC/x8UAv8fFAL/HxQC/x8UAv8fFAL/HxQC/x8UAv8fFAL/HxQC/x8UAv8fFAL/HxQC/x8UAv8bEQL/IRUC/4paDP+qbxD/qm8Q/6pvEP+qbxD/qm8Q/6pvEP+qbxD/qm8Q/6pvEP+qbxD/qm8Q/6pvEP+qbxD/OyYE/yIWAv+EVgz/mWMO/5ljDv+ZYw7/mWMO/5ljDv+ZYw7/mWMO/5ljDv+ZYw7/mWMO/5ljDv+ZYw7/mWMO/1c4B/8ZEAL/dk0K/6pvEP+qbxD/qm8Q/6pvEP+qbxD/qm8Q/6pvEP+qbxD/qm8Q/6pvEP+qbxD/qm8Q/6pvEP85JQT/JRgD/5NgDf+qbxD/qm8Q/6pvEP+qbxD/qm8Q/6pvEP+qbxD/qm8Q/6pvEP+qbxD/qm8Q/6pvEP+qbxD/Xj0I/xkQAv95Tgr/qm8Q/6pvEP+qbxD/qm8Q/6pvEP+qbxD/qm8Q/6pvEP+qbxD/qm8Q/6pvEP+qbxD/qm8Q/zklBP8lGAP/k2AN/6pvEP+qbxD/qm8Q/6pvEP+qbxD/qm8Q/6pvEP+qbxD/qm8Q/6pvEP+qbxD/qm8Q/6pvEP9ePQj/GRAC/3lOCv+qbxD/qm8Q/6pvEP+qbxD/qm8Q/6pvEP+qbxD/qm8Q/6pvEP+qbxD/qm8Q/6pvEP+qbxD/OSUE/yUYA/+TYA3/qm8Q/6pvEP+qbxD/qm8Q/6pvEP+qbxD/qm8Q/6pvEP+qbxD/qm8Q/6pvEP+qbxD/qm8Q/149CP8ZEAL/eU4K/6pvEP+qbxD/qm8Q/6pvEP+qbxD/qm8Q/6pvEP+qbxD/qm8Q/6pvEP+qbxD/qm8Q/6pvEP85JQT/JRgD/5NgDf+qbxD/qm8Q/6pvEP+qbxD/qm8Q/6pvEP+qbxD/qm8Q/6pvEP+qbxD/qm8Q/6pvEP+qbxD/Xj0I/xkQAv95Tgr/qm8Q/6pvEP+qbxD/qm8Q/6pvEP+qbxD/qm8Q/6pvEP+qbxD/qm8Q/6pvEP+qbxD/qm8Q/zklBP8lGAP/k2AN/6pvEP+qbxD/qm8Q/6pvEP+qbxD/qm8Q/6pvEP+qbxD/qm8Q/6pvEP+qbxD/qm8Q/6pvEP9ePQj/GRAC/3lOCv+qbxD/qm8Q/6pvEP+qbxD/qm8Q/6pvEP+qbxD/qm8Q/6pvEP+qbxD/qm8Q/6pvEP+qbxD/OSUE/yUYA/+TYA3/qm8Q/6pvEP+qbxD/qm8Q/6pvEP+qbxD/qm8Q/6pvEP+qbxD/qm8Q/6pvEP+qbxD/qm8Q/149CP8ZEAL/eU4K/6pvEP+qbxD/qm8Q/6pvEP+qbxD/qm8Q/6pvEP+qbxD/qm8Q/6pvEP+qbxD/qm8Q/6pvEP85JQT/JRgD/5NgDf+qbxD/qm8Q/6pvEP+qbxD/qm8Q/6pvEP+qbxD/qm8Q/6pvEP+qbxD/qm8Q/6pvEP+qbxD/Xj0I/xkQAv95Tgr/qm8Q/6pvEP+qbxD/qm8Q/6pvEP+qbxD/qm8Q/6pvEP+qbxD/qm8Q/6pvEP+qbxD/qm8Q/zklBP8lGAP/k2AN/6pvEP+qbxD/qm8Q/6pvEP+qbxD/qm8Q/6pvEP+qbxD/qm8Q/6pvEP+qbxD/qm8Q/6pvEP9ePQj/GRAC/3lOCv+qbxD/qm8Q/6pvEP+qbxD/qm8Q/6pvEP+qbxD/qm8Q/6pvEP+qbxD/qm8Q/6pvEP+qbxD/SC8G/xkQAv8qGwP/Lh4E/y4eBP8uHgT/Lh4E/y4eBP8uHgT/Lh4E/y4eBP8uHgT/Lh4E/y4eBP8uHgT/Lh4E/yIWAv8bEQL/hFYL/6pvEP+qbxD/qm8Q/6pvEP+qbxD/qm8Q/6pvEP+qbxD/qm8Q/6pvEP+qbxD/qm8Q/6pvEP+NWwz/PigF/y8fBP8vHwT/Lx8E/y8fBP8vHwT/Lx8E/y8fBP8vHwT/Lx8E/y8fBP8vHwT/Lx8E/y8fBP8vHwT/MB8E/1U4CP+kaw//qm8Q/6pvEP+qbxD/qm8Q/6pvEP+qbxD/qm8Q/6pvEP+qbxD/qm8Q/6pvEP+qbxD/qm8Q/6pvEP+lbA//nmcP/55nD/+eZw//nmcP/55nD/+eZw//nmcP/55nD/+eZw//nmcP/55nD/+eZw//nmcP/55nD/+gaA//qW4P/6pvEP+qbxD/qm8Q/6pvEP+qbxD/qm8Q/6pvEP+qbxD/qm8Q/6pvEP+qbxD/qm8Q/6pvEP+qbxD/qm8Q/6pvEP+qbxD/qm8Q/6pvEP+qbxD/qm8Q/6pvEP+qbxD/qm8Q/6pvEP+qbxD/qm8Q/6pvEP+qbxD/qm8Q/6pvEP+qbxD/qm8Q/6pvEP+qbxD/qm8Q/6pvEP+qbxD/qm8Q/6pvEP+qbxD/qm8Q/6pvEP+qbxD/qm8Q/6pvEP+qbxD/qm8Q/6pvEP+qbxD/qm8Q/6pvEP+qbxD/qm8Q/6pvEP+qbxD/qm8Q/6pvEP+qbxD/qm8Q/6pvEP+qbxD/qm8Q/6pvEP+qbxD/qm8Q/6pvEP+qbxD/qm8Q/6pvEP+qbxD/qm8Q/6pvEP+qbxD/qm8Q/6pvEP+qbxD/qm8Q/6pvEP+qbxD/qm8Q/6pvEP+qbxD/qm8Q/6pvEP+qbxD/qm8Q/6pvEP+qbxD/qm8Q/6pvEP+qbxD/qm8Q/6pvEP+qbxD/qm8Q/6pvEP+qbxD/qm8Q/6pvEP+qbxD/qm8Q/6pvEP+qbxD/qm8Q/6pvEP+qbxD/qm8Q/6pvEP+qbxD/qm8Q/6pvEP+qbxD/qm8Q/6pvEP+qbxD/qm8Q/6pvEP+qbxD/qm8Q/6pvEP+qbxD/qm8Q/6pvEP+qbxD/qm8Q/6pvEP+qbxD/qm8Q/6pvEP+qbxD/qm8Q/6pvEP+qbxD/qm8Q/6pvEP+qbxD/qm8Q/6pvEP+qbxD/qm8Q/6pvEP+qbxD/qm8Q/6pvEP+qbxD/qm8Q/6pvEP+qbxD/qm8Q/6pvEP+qbxD/qm8Q/6pvEP+qbxD/qm8Q/6pvEP+qbxD/qm8Q/6pvEP+qbxD/qm8Q/6pvEP+qbxD/qm8Q/6pvEP+qbxD/qm8Q/6pvEP+qbxD/qm8Q/6pvEP+qbxD/qm8Q/6pvEP+qbxD/qm8Q/6pvEP+qbxD/qm8Q/6pvEP+qbxD/qm8Q/6pvEP+qbxD/qm8Q/6pvEP+qbxD/qm8Q/6pvEP+qbxD/qm8Q/6pvEP+qbxD/qm8Q/6pvEP+qbxD/qm8Q/6pvEP+qbxD/qm8Q/6pvEP+qbxD/qm8Q/6pvEP+qbxD/qm8Q/6pvEP+qbxD/qm8Q/6pvEP+qbxD/qm8Q/6pvEP+qbxD/qm8Q/6pvEP+qbxD/qm8Q/6pvEP+qbxD/qm8Q/6pvEP+qbxD/qm8Q/6pvEP+qbxD/qm8Q/6pvEP+qbxD/qm8Q/6pvEP+qbxD/qm8Q/6pvEP+qbxD/qm8Q/6pvEP+qbxD/qm8Q/6pvEP+qbxD/qm8Q/6pvEP+qbxD/qm8Q/6pvEP+qbxD/qm8Q/6pvEP+qbxD/qm8Q/6pvEP+qbxD/qm8Q/6pvEP+qbxD/qm8Q/6pvEP+qbxD/qm8Q/6pvEP+qbxD/AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAoAAAAIAAAAEAAAAABAAgAAAAAAAAIAAAAAAAAAAAAAAAAAAAAAAAAFw8C/xkQAv8aEQL/GxIC/xwSAv8eFAP/IBUD/yIWA/8jFwP/JhkD/yobA/8rHAT/Lh4E/y8fBP8wHwT/MSAE/zklBf87JwX/PykG/0UtBv9JMAb/UTUH/1M2CP9WOAj/VzkI/109Cf9ePQj/bUcK/3dNC/95Twv/elAL/31RDP9+Ugz/hFYM/4VXDP+KWg3/jVwN/5BeDf+UYA7/l2MO/5lkD/+eZw//n2cP/6BpD/+kaw//pWsP/6ZsEP+pbxD/qm8Q/wAAAP8AAAD/AAAA/wAAAP8AAAD/AAAA/wAAAP8AAAD/AAAA/wAAAP8AAAD/AAAA/wAAAP8AAAD/AAAA/wAAAP8AAAD/AAAA/wAAAP8AAAD/AAAA/wAAAP8AAAD/AAAA/wAAAP8AAAD/AAAA/wAAAP8AAAD/AAAA/wAAAP8AAAD/AAAA/wAAAP8AAAD/AAAA/wAAAP8AAAD/AAAA/wAAAP8AAAD/AAAA/wAAAP8AAAD/AAAA/wAAAP8AAAD/AAAA/wAAAP8AAAD/AAAA/wAAAP8AAAD/AAAA/wAAAP8AAAD/AAAA/wAAAP8AAAD/AAAA/wAAAP8AAAD/AAAA/wAAAP8AAAD/AAAA/wAAAP8AAAD/AAAA/wAAAP8AAAD/AAAA/wAAAP8AAAD/AAAA/wAAAP8AAAD/AAAA/wAAAP8AAAD/AAAA/wAAAP8AAAD/AAAA/wAAAP8AAAD/AAAA/wAAAP8AAAD/AAAA/wAAAP8AAAD/AAAA/wAAAP8AAAD/AAAA/wAAAP8AAAD/AAAA/wAAAP8AAAD/AAAA/wAAAP8AAAD/AAAA/wAAAP8AAAD/AAAA/wAAAP8AAAD/AAAA/wAAAP8AAAD/AAAA/wAAAP8AAAD/AAAA/wAAAP8AAAD/AAAA/wAAAP8AAAD/AAAA/wAAAP8AAAD/AAAA/wAAAP8AAAD/AAAA/wAAAP8AAAD/AAAA/wAAAP8AAAD/AAAA/wAAAP8AAAD/AAAA/wAAAP8AAAD/AAAA/wAAAP8AAAD/AAAA/wAAAP8AAAD/AAAA/wAAAP8AAAD/AAAA/wAAAP8AAAD/AAAA/wAAAP8AAAD/AAAA/wAAAP8AAAD/AAAA/wAAAP8AAAD/AAAA/wAAAP8AAAD/AAAA/wAAAP8AAAD/AAAA/wAAAP8AAAD/AAAA/wAAAP8AAAD/AAAA/wAAAP8AAAD/AAAA/wAAAP8AAAD/AAAA/wAAAP8AAAD/AAAA/wAAAP8AAAD/AAAA/wAAAP8AAAD/AAAA/wAAAP8AAAD/AAAA/wAAAP8AAAD/AAAA/wAAAP8AAAD/AAAA/wAAAP8AAAD/AAAA/wAAAP8AAAD/AAAA/wAAAP8AAAD/AAAA/zAwMDAwMDAwMDAwMDAwMDAwMDAwMDAwMDAwMDAwMDAwMDAwMDAwMDAwMDAwMDAwMDAwMDAwMDAwMDAwMDAwMDAwMDAwMDAwMDAwMDAwMDAwMDAwMDAwMDAwMDAwMDAwMDAwMDAwMDAwMDAwMDAwMDAwMDAwMDAwMDAwMDAwMDAwMDAwMDAwMDAwMDAwMDAwMDAwMDAwMDAwMDAwMDAwMDAwMDAwMDAwMDAwMDAwMDAwMDAwMDAwMDAwMDAwMDAwMDAwMDAwMDAwMDAwMDAwMDAwMDAwMDAwMDAwMDAwMDAwMDAwMDAwMDAwMDAwMDAwMDAwMDAwMDAwMDAwMDAwMDAwMDAwMB8ZGRkZGRkZGRkZGRkZGRkZGRkZGRkrMDAwMDAwMDAwFgAAAAAAAAAAAAAAAAAAAAAAAAAAACcwMDAwMDAwMDAlIBsKAAAAAAAAAAAAAAAAAAABEx4gLDAwMDAwMDAwMDAwFQEFBgYGBgYGBgYGBgYGBgMHIzAwMDAwMDAwMDAwMDARCCEoKCgoKCgoKCgoKCgoGAEcMDAwMDAwMDAwMDAwMBAJJjAwMDAwMDAwMDAwMDAaAh0wMDAwMDAwMDAwMDAwEAkmMDAwMDAwMDAwMDAwMBoCHTAwMDAwMDAwMDAwMDAQCSYwMDAwMDAwMDAwMDAwGgIdMDAwMDAwMDAwMDAwMBAJJjAwMDAwMDAwMDAwMDAaAh0wMDAwMDAwMDAwMDAwEAkmMDAwMDAwMDAwMDAwMBoCHTAwMDAwMDAwMDAwMDAQCSYwMDAwMDAwMDAwMDAwGgIdMDAwMDAwMDAwMDAwMBAJJjAwMDAwMDAwMDAwMDAaAh0wMDAwMDAwMDAwMDAwEAkmMDAwMDAwMDAwMDAwMBoCHTAwMDAwMDAwMDAwMDAUAQsMDAwMDAwMDAwMDAwMBwQiMDAwMDAwMDAwMDAwMCQSDg0NDQ0NDQ0NDQ0NDQ0PFy0wMDAwMDAwMDAwMDAwMC4qKSkpKSkpKSkpKSkpKSsvMDAwMDAwMDAwMDAwMDAwMDAwMDAwMDAwMDAwMDAwMDAwMDAwMDAwMDAwMDAwMDAwMDAwMDAwMDAwMDAwMDAwMDAwMDAwMDAwMDAwMDAwMDAwMDAwMDAwMDAwMDAwMDAwMDAwMDAwMDAwMDAwMDAwMDAwMDAwMDAwMDAwMDAwMDAwMDAwMDAwMDAwMDAw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AP8AAAD/AAAA/wAAAP8AAAD/AAAA/wAAAP8AAAD/AAAA/wAAAP8AAAD/AAAA/wAAAP8AAAD/AAAA/wAAAP8AAAD/AAAA/wAAAP8AAAD/AAAA/wAAAP8AAAD/AAAA/wAAAP8AAAD/AAAA/wAAAP8AAAD/AAAA/wAAAP8AAAD/AAAA/wAAAP8AAAD/AAAA/wAAAP8AAAD/AAAA/wAAAP8AAAD/AAAA/wAAAP8AAAD/AAAA/wAAAP8AAAD/AAAA/wAAAP8AAAD/AAAA/wAAAP8AAAD/AAAA/wAAAP8AAAD/AAAA/wAAAP8AAAD/AAAA/wAAAP8AAAD/AAAA/wAAAP8AAAD/AAAA/wAAAP8AAAD/AAAA/wAAAP8AAAD/AAAA/wAAAP8AAAD/AAAA/wAAAP8AAAD/AAAA/wAAAP8AAAD/AAAA/wAAAP8AAAD/AAAA/wAAAP8AAAD/AAAA/wAAAP8AAAD/AAAA/wAAAP8AAAD/AAAA/wAAAP8AAAD/AAAA/wAAAP8AAAD/AAAA/wAAAP8AAAD/AAAA/wAAAP8AAAD/AAAA/wAAAP8AAAD/AAAA/wAAAP8AAAD/AAAA/wAAAP8AAAD/AAAA/wAAAP8AAAD/AAAA/wAAAP8AAAD/AAAA/wAAAP8AAAD/AAAA/wAAAP8AAAD/AAAA/wAAAP8AAAD/AAAA/wAAAP8AAAD/AAAA/wAAAP8AAAD/AAAA/wAAAP8AAAD/AAAA/wAAAP8AAAD/AAAA/wAAAP8AAAD/AAAA/wAAAP8AAAD/AAAA/wAAAP8AAAD/AAAA/wAAAP8AAAD/AAAA/wAAAP8AAAD/AAAA/wAAAP8AAAD/AAAA/wAAAP8AAAD/GxsbGxsbGxsbGxsbGxsbGxsbGxsbGxsbGxsbGxsbGxsbGxsbGxsbGxsbGxsbGxsbGxsbGxsbGxsbGxsbGxsbGxsbFgMDAwMDAwMDAwMOGxsbGxoTAgEBAQEBAQAKGBsbGxsbEQkZGRkZGRkUBRsbGxsbGxELGxsbGxsbFQYbGxsbGxsRCxsbGxsbGxUGGxsbGxsbEQsbGxsbGxsVBhsbGxsbGxIEDw8PDw8PCAcbGxsbGxsaEAwMDAwMDA0XGxsbGxsbGxsbGxsbGxsbGxsbGxsbGxsbGxsbGxsbGxsbGxsbGxsbGxsbGxsbGxsbGxsbGxsbGxsbGxsbGxsbGxsbGwAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA=")
        $NotifyIcon.BalloonTipIcon = [System.Windows.Forms.ToolTipIcon]::Info
        $NotifyIcon.BalloonTipTitle = $GUIText.BalloonTipTitle
        $NotifyIcon.BalloonTipText = $GUIText.BalloonTipText
        $NotifyIcon.Visible = $true

        # Set Window location to the bottom right corner
        Set-WindowWorkArea

        #Registrer event when icon is clicked
        [void](Register-ObjectEvent -InputObject $NotifyIcon -EventName MouseDoubleClick -SourceIdentifier IconClicked -Action {

            #Show window
            if($WPFWindow.Window.Visibility -eq "Hidden"){
                Set-WindowWorkArea
                $WPFWindow.Window.Visibility = "Visible"
            }
        }) 
    }else{

        #Disable hide button
        $WPFWindow.Window.WindowStartupLocation = "CenterScreen"
        $WPFWindow.btn_minimize.IsEnabled = $false
    }

    #Drag window functionality
    $WPFWindow.Window.Add_MouseLeftButtonDown({
        # $WPFWindow.Window.DragMove()
    })

    #Minimize click
    $WPFWindow.btn_minimize.Add_Click({

        #If the job is alredy done, exit
        if($CheckBatteryTimer.Enabled -eq $False){
                
            #Clear and close
            Clear-AndClose
        }

        #Hide Window
        $WPFWindow.Window.Visibility = "Hidden"

        #Show notification
        $NotifyIcon.ShowBalloonTip(5000)
    })

    #Exit click
    $WPFWindow.btn_exit.Add_Click({

        #Clear and close
        Clear-AndClose
    })

    #On closing
    $WPFWindow.Window.add_Closing({

        #Cancel if the user tries to close the window
        if($RunningTaskSequence -eq $false){
            $_.Cancel = $true
        }
    }) 

    #Start timer
    Start-CheckBattery

    #Set textboxes
    $WPFWindow.tbx_title.Text = $GUIText.TitleText
    $WPFWindow.tbx_header.Text = ($GUIText.HeaderText -f  $Model)
    $WPFWindow.tbx_body.Text = $GUIText.BodyText
    $WPFWindow.tbx_message.Text = ($GUIText.BottomText -f $Global:MinimumBatteryPercentage)
    
    #Make PowerShell Disappear 
    $WindowCode = '[DllImport("user32.dll")] public static extern bool ShowWindowAsync(IntPtr hWnd, int nCmdShow);' 
    $AsyncWindow = Add-Type -MemberDefinition $WindowCode -name Win32ShowWindowAsync -namespace Win32Functions -PassThru 
    $null = $AsyncWindow::ShowWindowAsync((Get-Process -PID $pid).MainWindowHandle, 0) 
    
    #Running this without $appContext and ::Run would actually cause a really poor response.
    $WPFWindow.Window.ShowDialog()

    #This makes it pop up
    $WPFWindow.Window.Activate()
    
    #Create an application context for it to all run within. 
    $AppContext = New-Object System.Windows.Forms.ApplicationContext
    [void][System.Windows.Forms.Application]::Run($AppContext)
}