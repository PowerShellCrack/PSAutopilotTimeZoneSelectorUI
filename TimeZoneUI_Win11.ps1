
<#
    .SYNOPSIS
        Prompts user to set time zone

    .DESCRIPTION
		Prompts user to set time zone using windows presentation framework design that looks like Windows 11 OOBE
        Can be used in:
            - SCCM Tasksequences (User interface allowed)
            - SCCM Software Delivery (User interface allowed)
            - Intune Autopilot

    .NOTES
        Author		: Dick Tracy <richard.tracy@hotmail.com>
	    Source		: https://github.com/PowerShellCrack/AutopilotTimeZoneSelectorUI
        Version		: 2.0.5
        README      : Review README.md for more details and configurations
        CHANGELOG   : Review CHANGELOG.md for updates and fixes
        IMPORTANT   : By using this script or parts of it, you have read and accepted the DISCLAIMER.md and LICENSE agreement

    .LINK
        https://matthewjwhite.co.uk/2019/04/18/intune-automatically-set-timezone-on-new-device-build/
        https://ipstack.com
        https://azuremarketplace.microsoft.com/en-us/marketplace/apps/bingmaps.mapapis

    .PARAMETER SyncNTP
        String --> defaults to 'pool.ntp.org'
        If value exist, the script will attempt to sync time with NTP.
        If this is not desired, remove the value or call it '-SynNTP $Null'
        NTP uses port UDP 123

    .PARAMETER IpStackAPIKey
        String --> value is Null
        Used to get geoCoordinates of the public IP. get the API key from https://ipstack.com

    .PARAMETER BingMapsAPIKeyy
        String --> value is Null
        Used to get the Windows TimeZone value of the location coordinates. get the API key from https://azuremarketplace.microsoft.com/en-us/marketplace/apps/bingmaps.mapapis

    .PARAMETER NoControl
        Boolean (True or False) --> Default is False
        Used for single deployment scenarios; This will not track status by writing registry keys.
        WARNING: UserDriven & RunOnce options are **IGNORED**.

    .PARAMETER UserDriven
        Boolean (True or False) --> Default is True
        Deploy to user when set to true.
        if _true_ sets HKCU key, if _false_, set HKLM key.
        Set to True if the deployment is for Autopilot.
        When using 'Users context deployment' and UserDriven is ste to False; users will need permission to write to device registry hive

    .PARAMETER NoUI
        Boolean (True or False) --> Default is False
        If set to True, the UI will not show but still attempt to set the timezone.
        If API Keys are provided it will use the internet to determine location.
        If Keys are not set, then it won't change the timezone because its the same as before, but it will attempt to sync time if a NTP value is provided.

    .PARAMETER RunOnce
        Boolean (True or False) --> Default is True
        Specifies this script will only launch one time.
        If RunOnce set to True and UserDriven is True, it will launch once for each user on the device.
        If RunOnce set to True and UserDriven is False, it will only launch once for the first user to login on the device
        If RunOnce set to False and UserDriven is True and on a reoccurring schedule, it will launch once for each user every time for each occupance (NOT RECOMMENDED)
        If RunOnce set to False and UserDriven is False and on a reoccurring schedule, it will launch for the current user logged in each occupance (NOT RECOMMENDED)

    .PARAMETER ForceInteraction
        Boolean (True or False) --> Default is False
        If set to True, no matter the other settings (including NoUI), the UI will **ALWAYS** show!


    .EXAMPLE
        NOTE: this example is using invalid API keys.
        PS> .\TimeZoneUI_Win11.ps1 -IpStackAPIKey "4bd1443445dfhrrt9dvefr45341" -BingMapsAPIKey "Bh53uNUOwg71czosmd73hKfdHf465ddfhrtpiohvknlkewufjf4-d" -Verbose

        RESULT: Uses IP GEO location for the pre-selection

    .EXAMPLE
        NOTE: this example is using invalid API keys.
        PS> .\TimeZoneUI_Win11.ps1 -ForceInteraction:$true -verbose

        RESULT:  This will ALWAYS display the time selection screen; if IPStack and BingMapsAPI included the IP GEO location timezone will be preselected. Verbose output will be displayed

    .EXAMPLE
        PS> .\TimeZoneUI_Win11.ps1 -IpStackAPIKey "4bd1443445dfhrrt9dvefr45341" -BingMapsAPIKey "Bh53uNUOwg71czosmd73hKfdHf465ddfhrtpiohvknlkewufjf4-d" -NoUI:$true -SyncNTP "time-a-g.nist.gov"

        RESULT: This will set the time automatically using the IP GEO location without prompting user. If API not provided, timezone or time will not change the current settings

    .EXAMPLE
        PS> .\TimeZoneUI_Win11.ps1 -UserDriven:$false

        RESULT: Writes a registry key in System (HKEY_LOCAL_MACHINE) hive to determine run status

    .EXAMPLE
        PS> .\TimeZoneUI_Win11.ps1 -RunOnce:$true

        RESULT: This allows the screen to display one time. RECOMMENDED for Autopilot to display after ESP screen
#>


#===========================================================================
# CONTROL VARIABLES
#===========================================================================

[CmdletBinding()]
param(

    [string]$SyncNTP = 'pool.ntp.org',

    [string]$IpStackAPIKey = "",

    [string]$BingMapsAPIKey = "" ,

    [boolean]$NoControl = $False,

    [boolean]$UserDriven = $true,

    [boolean]$RunOnce = $true,

    [boolean]$NoUI = $False,

    [boolean]$ForceInteraction = $false
)

#*=============================================
##* Runtime Function - REQUIRED
##*=============================================
#region FUNCTION: Check if running in WinPE
Function Test-WinPE{
    return Test-Path -Path Registry::HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlset\Control\MiniNT
}
#endregion

#region FUNCTION: Check if running in ISE
Function Test-IsISE {
    # try...catch accounts for:
    # Set-StrictMode -Version latest
    try {
        return ($null -ne $psISE);
    }
    catch {
        return $false;
    }
}
#endregion

#region FUNCTION: Check if running in Visual Studio Code
Function Test-VSCode{
    if($env:TERM_PROGRAM -eq 'vscode') {
        return $true;
    }
    Else{
        return $false;
    }
}
#endregion

#region FUNCTION: Find script path for either ISE or console
Function Get-ScriptPath {
    <#
        .SYNOPSIS
            Finds the current script path even in ISE or VSC
        .LINK
            Test-VSCode
            Test-IsISE
    #>
    param(
        [switch]$Parent
    )

    Begin{}
    Process{
        if ($PSScriptRoot -eq "")
        {
            if (Test-IsISE)
            {
                $ScriptPath = $psISE.CurrentFile.FullPath
            }
            elseif(Test-VSCode){
                $context = $psEditor.GetEditorContext()
                $ScriptPath = $context.CurrentFile.Path
            }Else{
                $ScriptPath = (Get-location).Path
            }
        }
        else
        {
            $ScriptPath = $PSCommandPath
        }
    }
    End{

        If($Parent){
            Split-Path $ScriptPath -Parent
        }Else{
            $ScriptPath
        }
    }

}
#endregion


#region FUNCTION: Attempt to connect to Task Sequence environment
Function Test-SMSTSENV{
    <#
        .SYNOPSIS
            Tries to establish Microsoft.SMS.TSEnvironment COM Object when running in a Task Sequence

        .REQUIRED
            Allows Set Task Sequence variables to be set

        .PARAMETER ReturnLogPath
            If specified, returns the log path, otherwise returns ts environment
    #>
    [CmdletBinding()]
    param(
        [switch]$ReturnLogPath
    )

    Begin{
        ## Get the name of this function
        [string]${CmdletName} = $MyInvocation.MyCommand
    }
    Process{
        try{
            # Create an object to access the task sequence environment
            $tsenv = New-Object -ComObject Microsoft.SMS.TSEnvironment
            #grab the progress UI
            $TSProgressUi = New-Object -ComObject Microsoft.SMS.TSProgressUI
            Write-Verbose ("Task Sequence environment detected!")
        }
        catch{

            Write-Verbose ("Task Sequence environment NOT detected.")
            #set variable to null
            $tsenv = $null
        }
        Finally{
            #set global Logpath
            if ($null -ne $tsenv)
            {
                # Convert all of the variables currently in the environment to PowerShell variables
                #$tsenv.GetVariables() | ForEach-Object { Set-Variable -Name "$_" -Value "$($tsenv.Value($_))" }

                # Query the environment to get an existing variable
                # Set a variable for the task sequence log path

                #Something like C:\WINDOWS\CCM\Logs\SMSTSLog
                [string]$LogPath = $tsenv.Value("_SMSTSLogPath")
                If($null -eq $LogPath){$LogPath = $env:Temp}
            }
            Else{
                $LogPath = $env:Temp
                $tsenv = $false
            }
        }
    }
    End{
        If($ReturnLogPath){
            return $LogPath
        }
        Else{
            return $tsenv
        }
    }
}
#endregion

Function Write-LogEntry{
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true,Position=0,ValueFromPipeline=$true,ValueFromPipelineByPropertyName=$true)]
        [ValidateNotNullOrEmpty()]
        [string]$Message,

        [Parameter(Mandatory=$false,Position=2)]
		[string]$Source,

        [parameter(Mandatory=$false)]
        [ValidateSet(0,1,2,3,4,5)]
        [int16]$Severity = 1,

        [parameter(Mandatory=$false, HelpMessage="Name of the log file that the entry will written to.")]
        [ValidateNotNullOrEmpty()]
        [string]$OutputLogFile = $Global:LogFilePath,

        [parameter(Mandatory=$false)]
        [switch]$Outhost
    )
    ## Get the name of this function
    #[string]${CmdletName} = $PSCmdlet.MyInvocation.MyCommand.Name
    if (-not $PSBoundParameters.ContainsKey('Verbose')) {
        $VerbosePreference = $PSCmdlet.SessionState.PSVariable.GetValue('VerbosePreference')
    }

    if (-not $PSBoundParameters.ContainsKey('Debug')) {
        $DebugPreference = $PSCmdlet.SessionState.PSVariable.GetValue('DebugPreference')
    }
    #get BIAS time
    [string]$LogTime = (Get-Date -Format 'HH:mm:ss.fff').ToString()
	[string]$LogDate = (Get-Date -Format 'MM-dd-yyyy').ToString()
	[int32]$script:LogTimeZoneBias = [timezone]::CurrentTimeZone.GetUtcOffset([datetime]::Now).TotalMinutes
	[string]$LogTimePlusBias = $LogTime + $script:LogTimeZoneBias

    #  Get the file name of the source script
    If($Source){
        $ScriptSource = $Source
    }
    Else{
        Try {
    	    If ($script:MyInvocation.Value.ScriptName) {
    		    [string]$ScriptSource = Split-Path -Path $script:MyInvocation.Value.ScriptName -Leaf -ErrorAction 'Stop'
    	    }
    	    Else {
    		    [string]$ScriptSource = Split-Path -Path $script:MyInvocation.MyCommand.Definition -Leaf -ErrorAction 'Stop'
    	    }
        }
        Catch {
    	    $ScriptSource = ''
        }
    }

    #if the severity is 4 or 5 make them 1; but output as verbose or debug respectfully.
    If($Severity -eq 4){$logSeverityAs=1}Else{$logSeverityAs=$Severity}
    If($Severity -eq 5){$logSeverityAs=1}Else{$logSeverityAs=$Severity}

    #generate CMTrace log format
    $LogFormat = "<![LOG[$Message]LOG]!>" + "<time=`"$LogTimePlusBias`" " + "date=`"$LogDate`" " + "component=`"$ScriptSource`" " + "context=`"$([Security.Principal.WindowsIdentity]::GetCurrent().Name)`" " + "type=`"$logSeverityAs`" " + "thread=`"$PID`" " + "file=`"$ScriptSource`">"

    # Add value to log file
    try {
        Out-File -InputObject $LogFormat -Append -NoClobber -Encoding Default -FilePath $OutputLogFile -ErrorAction Stop
    }
    catch {
        Write-Host ("[{0}] [{1}] :: Unable to append log entry to [{1}], error: {2}" -f $LogTimePlusBias,$ScriptSource,$OutputLogFile,$_.Exception.ErrorMessage) -ForegroundColor Red
    }

    #output the message to host
    If($Outhost)
    {
        If($Source){
            $OutputMsg = ("[{0}] [{1}] :: {2}" -f $LogTimePlusBias,$Source,$Message)
        }
        Else{
            $OutputMsg = ("[{0}] [{1}] :: {2}" -f $LogTimePlusBias,$ScriptSource,$Message)
        }

        Switch($Severity){
            0       {Write-Host $OutputMsg -ForegroundColor Green}
            1       {Write-Host $OutputMsg -ForegroundColor Gray}
            2       {Write-Host $OutputMsg -ForegroundColor Yellow}
            3       {Write-Host $OutputMsg -ForegroundColor Red}
            4       {Write-Verbose $OutputMsg}
            5       {Write-Debug $OutputMsg}
            default {Write-Host $OutputMsg}
        }
    }
}
##*=============================================
##* VARIABLE DECLARATION
##*=============================================
#region VARIABLES: Building paths & values
[string]$scriptPath = Get-ScriptPath
[string]$scriptName = [IO.Path]::GetFileNameWithoutExtension($scriptPath)

#Grab all times zones plus current timezones
$Global:AllTimeZones = Get-TimeZone -ListAvailable
$Global:CurrentTimeZone = Get-TimeZone

$Global:NTPServer = $SyncNTP

#set the appropriate registry hive to use when logging
If($UserDriven -eq $false){$RegHive = 'HKLM'}Else{$RegHive = 'HKCU'}

#Return log path (either in task sequence or temp dir)
#build log name
[string]$FileName = $scriptName +'.log'
#build global log fullpath
$Global:LogFilePath = Join-Path (Test-SMSTSENV -ReturnLogPath -Verbose) -ChildPath $FileName
Write-Host "logging to file: $LogFilePath" -ForegroundColor Cyan
#===========================================================================
# XAML LANGUAGE
#===========================================================================
$XAML = @"
<Window x:Class="SelectTimeZoneWPF.MainWindow"
        xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        xmlns:d="http://schemas.microsoft.com/expression/blend/2008"
        xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
        xmlns:local="clr-namespace:SelectTimeZoneWPF"
        mc:Ignorable="d"
        WindowState="Maximized"
        WindowStartupLocation="CenterScreen"
        WindowStyle="None"
        Width="1024" Height="768"
        Title="Time Zone Selection"
        >
    <Window.Background>
        <ImageBrush ImageSource="https://github.com/PowerShellCrack/AutopilotTimeZoneSelectorUI/blob/master/.images/win11_oobe_wallpaper.png?raw=true"></ImageBrush>
    </Window.Background>
    <Window.Resources>
        <ResourceDictionary>

            <Style TargetType="{x:Type Window}">
                <Setter Property="FontFamily" Value="Segoe UI" />
                <Setter Property="FontWeight" Value="Normal" />
                <Setter Property="Background" Value="white" />
                <Setter Property="Foreground" Value="#1f1f1f" />
            </Style>

            <!-- TabControl Style-->
            <Style  TargetType="TabControl">
                <Setter Property="OverridesDefaultStyle" Value="true"/>
                <Setter Property="SnapsToDevicePixels" Value="true"/>
                <Setter Property="Template">
                    <Setter.Value>
                        <ControlTemplate TargetType="{x:Type TabControl}">
                            <Grid KeyboardNavigation.TabNavigation="Local">
                                <Grid.RowDefinitions>
                                    <RowDefinition Height="Auto" />
                                    <RowDefinition Height="*" />
                                </Grid.RowDefinitions>

                                <TabPanel x:Name="HeaderPanel"
                                  Grid.Row="0"
                                  Panel.ZIndex="1"
                                  Margin="0,0,4,-3"
                                  IsItemsHost="True"
                                  KeyboardNavigation.TabIndex="1"
                                  Background="Transparent" />

                                <Border x:Name="Border"
                            Grid.Row="1"
                                        CornerRadius="15"
                            BorderThickness="0,3,0,0"
                            KeyboardNavigation.TabNavigation="Local"
                            KeyboardNavigation.DirectionalNavigation="Contained"
                            KeyboardNavigation.TabIndex="2">

                                    <Border.Background>
                                        <SolidColorBrush Color="White" Opacity="0.7"/>
                                    </Border.Background>

                                    <Border.BorderBrush>
                                        <SolidColorBrush Color="#eef2f4" />
                                    </Border.BorderBrush>

                                    <Border.Effect>
                                        <DropShadowEffect BlurRadius="100" Color="#FFE3E3E3" ShadowDepth="0" />
                                    </Border.Effect>

                                    <ContentPresenter x:Name="PART_SelectedContentHost"
                                          Margin="0,0,0,0"
                                          ContentSource="SelectedContent" />
                                </Border>
                            </Grid>
                        </ControlTemplate>
                    </Setter.Value>
                </Setter>
            </Style>

            <!-- TabItem Style -->
            <Style x:Key="OOBETabStyle" TargetType="{x:Type TabItem}" >
                <!--<Setter Property="Foreground" Value="#FFE6E6E6"/>-->
                <Setter Property="Template">
                    <Setter.Value>

                        <ControlTemplate TargetType="{x:Type TabItem}">
                            <Grid>
                                <Border
                                    Name="Border"
                                    Margin="0"
                                    CornerRadius="0">
                                    <ContentPresenter x:Name="ContentSite" VerticalAlignment="Center"
                                        HorizontalAlignment="Center" ContentSource="Header"
                                        RecognizesAccessKey="True" />
                                </Border>
                            </Grid>

                            <ControlTemplate.Triggers>
                                <Trigger Property="IsMouseOver" Value="True">
                                    <Setter Property="Foreground" Value="#313131" />
                                    <Setter TargetName="Border" Property="BorderThickness" Value="0,0,0,3" />
                                    <Setter TargetName="Border" Property="BorderBrush" Value="#4c4c4c" />
                                </Trigger>
                                <Trigger Property="IsMouseOver" Value="False">
                                    <Setter Property="Foreground" Value="#313131" />
                                    <Setter TargetName="Border" Property="BorderThickness" Value="0,0,0,3" />
                                    <Setter TargetName="Border" Property="BorderBrush" Value="#4c4c4c" />
                                </Trigger>
                                <Trigger Property="IsSelected" Value="True">
                                    <Setter Property="Panel.ZIndex" Value="100" />
                                    <Setter Property="Foreground" Value="white" />
                                    <Setter TargetName="Border" Property="BorderThickness" Value="0,0,0,3" />
                                    <Setter TargetName="Border" Property="BorderBrush" Value="White" />
                                </Trigger>
                            </ControlTemplate.Triggers>
                        </ControlTemplate>
                    </Setter.Value>
                </Setter>

            </Style>

            <Style x:Key="DataGridContentCellCentering" TargetType="{x:Type DataGridCell}">
                <Setter Property="Template">
                    <Setter.Value>
                        <ControlTemplate TargetType="{x:Type DataGridCell}">
                            <Grid Background="{TemplateBinding Background}">
                                <ContentPresenter VerticalAlignment="Center" />
                            </Grid>
                        </ControlTemplate>
                    </Setter.Value>
                </Setter>
            </Style>

            <!-- Sub TabItem Style -->
            <!-- TabControl Style-->
            <Style x:Key="ModernStyleTabControl" TargetType="TabControl">
                <Setter Property="OverridesDefaultStyle" Value="true"/>
                <Setter Property="SnapsToDevicePixels" Value="true"/>
                <Setter Property="Template">

                    <Setter.Value>
                        <ControlTemplate TargetType="{x:Type TabControl}">
                            <Grid KeyboardNavigation.TabNavigation="Local">
                                <Grid.RowDefinitions>
                                    <RowDefinition Height="40" />
                                    <RowDefinition Height="*" />
                                </Grid.RowDefinitions>

                                <TabPanel x:Name="HeaderPanel"
                                    Grid.Row="0"
                                    Panel.ZIndex="1"
                                    IsItemsHost="True"
                                    KeyboardNavigation.TabIndex="1"
                                    Background="#eef2f4"  />

                                <Border x:Name="Border"
                                    Grid.Row="0"
                                    BorderThickness="1"
                                    BorderBrush="Black"
                                    Background="#eef2f4">

                                    <ContentPresenter x:Name="PART_SelectedContentHost"
                                          Margin="0,0,0,0"
                                          ContentSource="SelectedContent" />
                                </Border>
                                <Border Grid.Row="1"
                                        BorderThickness="1,0,1,1"
                                        BorderBrush="#eef2f4">
                                    <ContentPresenter Margin="4" />
                                </Border>
                            </Grid>
                        </ControlTemplate>
                    </Setter.Value>
                </Setter>
            </Style>


            <Style x:Key="ModernStyleTabItem" TargetType="{x:Type TabItem}">
                <Setter Property="Template">
                    <Setter.Value>

                        <ControlTemplate TargetType="{x:Type TabItem}">
                            <Grid>
                                <Border
                                    Name="Border"
                                    Margin="10,10,10,10"
                                    CornerRadius="5">
                                    <ContentPresenter x:Name="ContentSite" VerticalAlignment="Center"
                                        HorizontalAlignment="Center" ContentSource="Header"
                                        RecognizesAccessKey="True" />
                                </Border>
                            </Grid>

                            <ControlTemplate.Triggers>
                                <Trigger Property="IsMouseOver" Value="True">
                                    <Setter Property="Foreground" Value="#FF9C9C9C" />
                                    <Setter Property="FontSize" Value="16" />
                                    <Setter TargetName="Border" Property="BorderThickness" Value="1,0,1,1" />
                                    <Setter TargetName="Border" Property="BorderBrush" Value="#eef2f4" />
                                </Trigger>
                                <Trigger Property="IsMouseOver" Value="False">
                                    <Setter Property="Foreground" Value="#FF666666" />
                                    <Setter Property="FontSize" Value="16" />
                                    <Setter TargetName="Border" Property="BorderThickness" Value="1,0,1,1" />
                                    <Setter TargetName="Border" Property="BorderBrush" Value="#eef2f4" />
                                </Trigger>
                                <Trigger Property="IsSelected" Value="True">
                                    <Setter Property="Panel.ZIndex" Value="100" />
                                    <Setter Property="Foreground" Value="white" />
                                    <Setter Property="FontSize" Value="16" />
                                    <Setter TargetName="Border" Property="BorderThickness" Value="1,0,1,1" />
                                    <Setter TargetName="Border" Property="BorderBrush" Value="#eef2f4" />
                                </Trigger>
                            </ControlTemplate.Triggers>
                        </ControlTemplate>
                    </Setter.Value>
                </Setter>
            </Style>

            <Style TargetType="{x:Type Button}">
                <Setter Property="Background" Value="#0067c0" />
                <Setter Property="Foreground" Value="#FFE8EDF9" />
                <Setter Property="FontSize" Value="15" />
                <Setter Property="SnapsToDevicePixels" Value="True" />

                <Setter Property="Template">
                    <Setter.Value>
                        <ControlTemplate TargetType="Button" >

                            <Border Name="border"
                                BorderThickness="1"
                                Padding="4,2"
                                BorderBrush="#336891"
                                CornerRadius="8"
                                Background="#0078d7">
                                <ContentPresenter HorizontalAlignment="Center"
                                                VerticalAlignment="Center"
                                                TextBlock.TextAlignment="Center"
                                                />
                            </Border>

                            <ControlTemplate.Triggers>
                                <Trigger Property="IsMouseOver" Value="True">
                                    <Setter TargetName="border" Property="BorderBrush" Value="#FFE8EDF9" />
                                    <Setter Property="Background" Value="#FFE8EDF9" />
                                </Trigger>

                                <Trigger Property="IsPressed" Value="True">
                                    <Setter TargetName="border" Property="BorderBrush" Value="#003e92" />
                                    <Setter Property="Background" Value="#003e92" />
                                    <Setter Property="Effect">
                                        <Setter.Value>
                                            <DropShadowEffect ShadowDepth="0" Color="#003e92" Opacity="1" BlurRadius="10"/>
                                        </Setter.Value>
                                    </Setter>
                                </Trigger>
                                <Trigger Property="IsEnabled" Value="False">
                                    <Setter TargetName="border" Property="BorderBrush" Value="#dddfe1" />
                                    <Setter Property="Background" Value="#dddfe1" />
                                </Trigger>
                                <Trigger Property="IsFocused" Value="False">
                                    <Setter TargetName="border" Property="BorderBrush" Value="#dddfe1" />
                                    <Setter Property="Background" Value="#dddfe1" />
                                </Trigger>

                            </ControlTemplate.Triggers>
                        </ControlTemplate>
                    </Setter.Value>
                </Setter>
            </Style>

            <Style x:Key="{x:Type ListBox}" TargetType="ListBox">
                <Setter Property="Template">
                    <Setter.Value>
                        <ControlTemplate TargetType="ListBox">
                            <Border Name="Border" BorderThickness="1">
                                <ScrollViewer Margin="0" Focusable="false">
                                    <StackPanel Margin="2" IsItemsHost="True" />
                                </ScrollViewer>
                            </Border>
                        </ControlTemplate>
                    </Setter.Value>
                </Setter>
            </Style>

            <Style x:Key="{x:Type ListBoxItem}" TargetType="ListBoxItem">
                <Setter Property="ScrollViewer.CanContentScroll" Value="true" />
                <Setter Property="Template">
                    <Setter.Value>
                        <ControlTemplate TargetType="ListBoxItem">
                            <Border Name="ItemBorder" Padding="8" Margin="1" Background="#eef2f4" CornerRadius="8">
                                <ContentPresenter />
                            </Border>
                            <ControlTemplate.Triggers>
                                <Trigger Property="IsSelected" Value="True">
                                    <Setter TargetName="ItemBorder" Property="Background" Value="#003e92"/>
                                    <Setter Property="Foreground" Value="#eef2f4" />
                                </Trigger>
                                <MultiTrigger>
                                    <MultiTrigger.Conditions>
                                        <Condition Property="IsMouseOver" Value="True" />
                                        <Condition Property="IsSelected" Value="False" />
                                    </MultiTrigger.Conditions>
                                    <Setter TargetName="ItemBorder" Property="Background" Value="#dddfe1" />
                                </MultiTrigger>
                            </ControlTemplate.Triggers>
                        </ControlTemplate>
                    </Setter.Value>
                </Setter>
            </Style>

            <Style x:Key="CheckBoxModernStyle1" TargetType="{x:Type CheckBox}">
                <Setter Property="Foreground" Value="{DynamicResource {x:Static SystemColors.WindowTextBrushKey}}"/>
                <Setter Property="Background" Value="{DynamicResource {x:Static SystemColors.WindowBrushKey}}"/>
                <Setter Property="Template">
                    <Setter.Value>
                        <ControlTemplate TargetType="{x:Type CheckBox}">
                            <ControlTemplate.Resources>
                                <Storyboard x:Key="OnChecking">
                                    <DoubleAnimationUsingKeyFrames BeginTime="00:00:00" Storyboard.TargetName="slider" Storyboard.TargetProperty="(UIElement.RenderTransform).(TransformGroup.Children)[3].(TranslateTransform.X)">
                                        <SplineDoubleKeyFrame KeyTime="00:00:00.3000000" Value="32"/>
                                    </DoubleAnimationUsingKeyFrames>
                                </Storyboard>
                                <Storyboard x:Key="OnUnchecking">
                                    <DoubleAnimationUsingKeyFrames BeginTime="00:00:00" Storyboard.TargetName="slider" Storyboard.TargetProperty="(UIElement.RenderTransform).(TransformGroup.Children)[3].(TranslateTransform.X)">
                                        <SplineDoubleKeyFrame KeyTime="00:00:00.3000000" Value="0"/>
                                    </DoubleAnimationUsingKeyFrames>
                                    <ThicknessAnimationUsingKeyFrames BeginTime="00:00:00" Storyboard.TargetName="slider" Storyboard.TargetProperty="(FrameworkElement.Margin)">
                                        <SplineThicknessKeyFrame KeyTime="00:00:00.3000000" Value="1,1,1,1"/>
                                    </ThicknessAnimationUsingKeyFrames>
                                </Storyboard>
                            </ControlTemplate.Resources>

                            <DockPanel x:Name="dockPanel">
                                <ContentPresenter SnapsToDevicePixels="{TemplateBinding SnapsToDevicePixels}" Content="{TemplateBinding Content}" ContentStringFormat="{TemplateBinding ContentStringFormat}" ContentTemplate="{TemplateBinding ContentTemplate}" RecognizesAccessKey="True" VerticalAlignment="Center"/>
                                <Grid Margin="5,5,0,5" Width="66" Background="#eef2f4">

                                    <TextBlock Text="Yes" TextWrapping="Wrap" FontWeight="Bold" FontSize="18" HorizontalAlignment="Right" Margin="0,0,3,0" Foreground="White" VerticalAlignment="Center"/>
                                    <TextBlock Text="No"  TextWrapping="Wrap" FontWeight="Bold" FontSize="18" HorizontalAlignment="Left" Margin="2,0,0,0" Foreground="White" VerticalAlignment="Center"/>
                                    <Border HorizontalAlignment="Left" x:Name="slider" Width="32" BorderThickness="1,1,1,1" CornerRadius="0,0,0,0" RenderTransformOrigin="0.5,0.5" Margin="1,1,1,1" Height="40">
                                        <Border.RenderTransform>
                                            <TransformGroup>
                                                <ScaleTransform ScaleX="1" ScaleY="1"/>
                                                <SkewTransform AngleX="0" AngleY="0"/>
                                                <RotateTransform Angle="0"/>
                                                <TranslateTransform X="0" Y="0"/>
                                            </TransformGroup>
                                        </Border.RenderTransform>
                                        <Border.BorderBrush>
                                            <LinearGradientBrush EndPoint="0.5,1" StartPoint="0.5,0">
                                                <GradientStop Color="#FFFFFFFF" Offset="0"/>
                                                <GradientStop Color="#FF4490FF" Offset="1"/>
                                            </LinearGradientBrush>
                                        </Border.BorderBrush>
                                        <Border.Background>
                                            <LinearGradientBrush EndPoint="0.5,1" StartPoint="0.5,0">
                                                <GradientStop Color="#FF8AB4FF" Offset="1"/>
                                                <GradientStop Color="#FFD1E2FF" Offset="0"/>
                                            </LinearGradientBrush>
                                        </Border.Background>
                                    </Border>
                                </Grid>
                            </DockPanel>

                            <ControlTemplate.Triggers>
                                <Trigger Property="IsChecked" Value="True">
                                    <Trigger.ExitActions>
                                        <BeginStoryboard Storyboard="{StaticResource OnUnchecking}" x:Name="OnUnchecking_BeginStoryboard"/>
                                    </Trigger.ExitActions>
                                    <Trigger.EnterActions>
                                        <BeginStoryboard Storyboard="{StaticResource OnChecking}" x:Name="OnChecking_BeginStoryboard"/>
                                    </Trigger.EnterActions>
                                </Trigger>
                                <Trigger Property="IsEnabled" Value="False">
                                    <Setter Property="Foreground" Value="{DynamicResource {x:Static SystemColors.GrayTextBrushKey}}"/>
                                </Trigger>
                            </ControlTemplate.Triggers>
                        </ControlTemplate>
                    </Setter.Value>
                </Setter>
            </Style>
            <Style x:Key="ModernStyleGroupBox" TargetType="{x:Type GroupBox}">
                <Setter Property="Template">
                    <Setter.Value>
                        <ControlTemplate TargetType="GroupBox">
                            <Grid>
                                <Grid.RowDefinitions>
                                    <RowDefinition Height="Auto" />
                                    <RowDefinition Height="*" />
                                </Grid.RowDefinitions>

                                <Border Grid.Row="0"
                                        BorderThickness="1"
                                        BorderBrush="Black"
                                        Background="#FF1D3245">
                                    <Label Foreground="White">
                                        <ContentPresenter Margin="4"
                                                          ContentSource="Header"
                                                          RecognizesAccessKey="True" />
                                    </Label>
                                </Border>
                                <Border Grid.Row="1"
                                        BorderThickness="1,0,1,1"
                                        BorderBrush="#FF1D3245">
                                    <ContentPresenter Margin="4" />
                                </Border>
                            </Grid>
                        </ControlTemplate>
                    </Setter.Value>
                </Setter>
            </Style>
            <Style x:Key="ModernToggleButton" TargetType="{x:Type ToggleButton}">
                <Setter Property="MinWidth" Value="80"/>
                <Setter Property="MinHeight" Value="26"/>
                <Setter Property="Margin" Value="0"/>
                <Setter Property="Background" Value="#336891" />
                <Setter Property="Foreground" Value="White"/>
                <Style.Triggers>
                    <Trigger Property="IsMouseOver" Value="True">
                        <Setter Property="BorderBrush" Value="#FF1D3245" />
                        <Setter Property="Foreground" Value="#FF1D3245" />
                    </Trigger>
                    <Trigger Property="IsFocused" Value="True">
                        <Setter Property="BorderBrush" Value="#FF1D3245" />
                        <Setter Property="Foreground" Value="#FF1D3245" />
                    </Trigger>
                    <Trigger Property="IsChecked" Value="True">
                        <Setter Property="BorderBrush" Value="#FF1D3245" />
                        <Setter Property="Foreground" Value="#FF1D3245" />
                        <Setter Property="FontWeight" Value="Bold" />
                        <Setter Property="Effect">
                            <Setter.Value>
                                <DropShadowEffect ShadowDepth="0" Color="#FF1D3245" Opacity="1" BlurRadius="10"/>
                            </Setter.Value>
                        </Setter>
                    </Trigger>
                </Style.Triggers>
            </Style>
        </ResourceDictionary>
    </Window.Resources>

    <Grid HorizontalAlignment="Center" VerticalAlignment="Center">
        <Grid.RowDefinitions>
            <RowDefinition/>
            <RowDefinition Height="0*"/>
        </Grid.RowDefinitions>

        <TabControl HorizontalAlignment="Center" VerticalAlignment="Center" Width="900" Height="650" Margin="0,0,0,40">

            <TabItem Style="{DynamicResource OOBETabStyle}" Header="Time Zone" Width="167" Height="60" BorderThickness="0" Margin="357,658,-357,-658">
                <Grid Margin="10,6,7,9">
                    <Image x:Name="tz_world" HorizontalAlignment="Left" Height="134" VerticalAlignment="Center" Width="148" Margin="170,217,0,224" Source="https://github.com/PowerShellCrack/AutopilotTimeZoneSelectorUI/blob/master/.images/win11_tz_worldclock.png?raw=true"/>
                    <TextBlock x:Name="tab3Version" HorizontalAlignment="Right" VerticalAlignment="Top" FontSize="12" FontFamily="Segoe UI Light" Width="883" TextAlignment="right" Foreground="gray"/>

                    <TextBlock x:Name="txtTimeZoneTitle" HorizontalAlignment="Center" Text="@anchor" VerticalAlignment="Top" FontSize="24" Margin="457,34,10,0" Width="416" TextAlignment="Left" FontFamily="Segoe UI" Foreground="Black" TextWrapping="Wrap"/>

                    <ListBox x:Name="lbxTimeZoneList" HorizontalAlignment="Center" VerticalAlignment="Top" Background="#eef2f4" Foreground="Black" FontSize="16" Width="415" Height="394" Margin="440,90,28,0" ScrollViewer.VerticalScrollBarVisibility="Auto" SelectionMode="Single"/>
                    <Button x:Name="btnTZSelect" Content="Select" Height="45" Width="140" HorizontalAlignment="Right" VerticalAlignment="Bottom" FontSize="18" Padding="10" Margin="0,0,28,29"/>

                </Grid>
            </TabItem>

        </TabControl>
    </Grid>
</Window>
"@

#replace some default attributes to support powershell
[string]$XAML = $XAML -replace 'mc:Ignorable="d"','' -replace "x:N",'N'  -replace '^<Win.*', '<Window'

#=======================================================
# LOAD ASSEMBLIES
#=======================================================
[System.Reflection.Assembly]::LoadWithPartialName('System.Windows.Forms')  | out-null #creating Windows-based applications
[System.Reflection.Assembly]::LoadWithPartialName('WindowsFormsIntegration')  | out-null # Call the EnableModelessKeyboardInterop; allows a Windows Forms control on a WPF page.
If(Test-WinPE -or Test-IsISE){[System.Reflection.Assembly]::LoadWithPartialName('System.Windows.Application')  | out-null} #Encapsulates a Windows Presentation Foundation application.
[System.Reflection.Assembly]::LoadWithPartialName('System.ComponentModel') | out-null #systems components and controls and convertors
[System.Reflection.Assembly]::LoadWithPartialName('System.Data')           | out-null #represent the ADO.NET architecture; allows multiple data sources
[System.Reflection.Assembly]::LoadWithPartialName('presentationframework') | out-null #required for WPF
[System.Reflection.Assembly]::LoadWithPartialName('PresentationCore')      | out-null #required for WPF

#convert to XML
[xml]$XAML = $XAML
#Read XAML
$reader=(New-Object System.Xml.XmlNodeReader $xaml)
try{$TZSelectUI=[Windows.Markup.XamlReader]::Load( $reader )}
catch{
    Write-LogEntry ("Unable to load Windows.Markup.XamlReader. {0}" -f $_.Exception.Message) -Severity 3 -Outhost
    Exit $_.Exception.HResult
}

#===========================================================================
# Store Form Objects In PowerShell
#===========================================================================
#take the xaml properties and make them variables
$xaml.SelectNodes("//*[@Name]") | %{Set-Variable -Name "ui_$($_.Name)" -Value $TZSelectUI.FindName($_.Name)}

Function Get-FormVariables{
    if ($global:ReadmeDisplay -ne $true){
        Write-Verbose "To reference this display again, run Get-FormVariables"
        $global:ReadmeDisplay=$true
    }
    Write-Verbose "Displaying elements from the form"
    Get-Variable ui_*
}

If($DebugPreference){Get-FormVariables}

#====================
#Form Functions
#====================

Function Set-StatusKey{
    param(
        [parameter(Mandatory=$False)]
        [ValidateSet('HKLM','HKCU')]
        [string]$Hive = 'HKCU',
        [parameter(Mandatory=$True)]
        [string]$Name,
        [parameter(Mandatory=$True)]
        [string]$Value
    )
    Begin
    {
        If(!(Test-Path "$($Hive):\SOFTWARE\PowerShellCrack\TimeZoneSelector") ){
            New-Item -Path "$($Hive):\SOFTWARE" -Name "PowerShellCrack" -ErrorAction SilentlyContinue | Out-Null
            New-Item -Path "$($Hive):\SOFTWARE\PowerShellCrack" -Name 'TimeZoneSelector' -ErrorAction SilentlyContinue | Out-Null
        }
    }
    Process
    {
        Try{
            Set-ItemProperty -Path "$($Hive):\SOFTWARE\PowerShellCrack\TimeZoneSelector" -Name $Name -Value $Value -Force -ErrorAction Stop | Out-Null
        }
        Catch{
            Write-LogEntry ("Unable to set status key name [{0}] with value [{1}]. {2}" -f $Name,$Value,$_.Exception.Message) -Severity 3 -Outhost
        }
    }
    End
    {
        Set-ItemProperty -Path "$($Hive):\SOFTWARE\PowerShellCrack\TimeZoneSelector" -Name "LastRan" -Value (Get-Date) -Force -ErrorAction SilentlyContinue | Out-Null
    }
}

Function Start-TimeSelectorUI{
    <#TEST VALUES
    $UIObject=$TZSelectUI
    $UpdateStatusKeyHive=$RegHive
    $UpdateStatusKeyHive="HKLM"
    #>
    [CmdletBinding()]
    param(
        $UIObject,
        [string]$UpdateStatusKeyHive
    )
    If($PSBoundParameters.ContainsKey('UpdateStatusKeyHive')){Set-StatusKey -Hive $UpdateStatusKeyHive -Name Status -Value "Running"}

    Try{
        #$UIObject.ShowDialog() | Out-Null
        # Credits to - http://powershell.cz/2013/04/04/hide-and-show-console-window-from-gui/
        Add-Type -Name Window -Namespace Console -MemberDefinition '
[DllImport("Kernel32.dll")]
public static extern IntPtr GetConsoleWindow();

[DllImport("user32.dll")]
public static extern bool ShowWindow(IntPtr hWnd, Int32 nCmdShow);
'
        # Allow input to window for TextBoxes, etc
        [Void][System.Windows.Forms.Integration.ElementHost]::EnableModelessKeyboardInterop($UIObject)

        #for ISE testing only: Add ESC key as a way to exit UI
        $code = {
            [System.Windows.Input.KeyEventArgs]$esc = $args[1]
            if ($esc.Key -eq 'ESC')
            {
                $UIObject.Close()
                [System.Windows.Forms.Application]::Exit()
                #this will kill ISE
                [Environment]::Exit($ExitCode);
            }
        }
        $null = $UIObject.add_KeyUp($code)

        $UIObject.Add_Closing({
            [System.Windows.Forms.Application]::Exit()
        })

        $async = $UIObject.Dispatcher.InvokeAsync({
            #make sure this display on top of every window
            $UIObject.Topmost = $true
            # Running this without $appContext & ::Run would actually cause a really poor response.
            $UIObject.Show() | Out-Null
            # This makes it pop up
            $UIObject.Activate() | Out-Null

            #$UI.window.ShowDialog()
        })
        $async.Wait() | Out-Null

        ## Force garbage collection to start form with slightly lower RAM usage.
        [System.GC]::Collect() | Out-Null
        [System.GC]::WaitForPendingFinalizers() | Out-Null

        # Create an application context for it to all run within.
        # This helps with responsiveness, especially when Exiting.
        $appContext = New-Object System.Windows.Forms.ApplicationContext
        [void][System.Windows.Forms.Application]::Run($appContext)
    }
    Catch{
        If($PSBoundParameters.ContainsKey('UpdateStatusKeyHive')){Set-StatusKey -Hive $UpdateStatusKeyHive -Name Status -Value 'Failed'}
        Write-LogEntry ("Unable to load Windows Presentation UI. {0}" -f $_.Exception.Message) -Severity 3 -Outhost
        Exit $_.Exception.HResult
    }
}

function Stop-TimeSelectorUI{
    <#TEST VALUES
    $UIObject=$TZSelectUI
    $UpdateStatusKeyHive="$RegHive\$RegPath"
    $UpdateStatusKeyHive="HKLM:\SOFTWARE\PowerShellCrack\TimeZoneSelector"
    #>
    [CmdletBinding()]
    param(
        $UIObject,
        [string]$UpdateStatusKeyHive,
        [string]$CustomStatus
    )

    If($CustomStatus){$status = $CustomStatus}
    Else{$status = 'Completed'}

    Try{
        If($PSBoundParameters.ContainsKey('UpdateStatusKeyHive')){Set-StatusKey -Hive $UpdateStatusKeyHive -Name Status -Value $status}
        #$UIObject.Close() | Out-Null
        $UIObject.Close()
    }
    Catch{
        If($PSBoundParameters.ContainsKey('UpdateStatusKeyHive')){Set-StatusKey -Hive $UpdateStatusKeyHive -Name Status -Value 'Failed'}
        Write-LogEntry ("Failed to stop Windows Presentation UI properly. {0}" -f $_.Exception.Message) -Severity 2 -Outhost
        #Exit $_.Exception.HResult
    }
}


function Set-NTPDateTime
{
    [CmdletBinding()]
    param(
        [string] $sNTPServer
    )

    $StartOfEpoch=New-Object DateTime(1900,1,1,0,0,0,[DateTimeKind]::Utc)
    [Byte[]]$NtpData = ,0 * 48
    $NtpData[0] = 0x1B    # NTP Request header in first byte
    $Socket = New-Object Net.Sockets.Socket([Net.Sockets.AddressFamily]::InterNetwork, [Net.Sockets.SocketType]::Dgram, [Net.Sockets.ProtocolType]::Udp)
    $Socket.Connect($sNTPServer,123)

    $t1 = Get-Date    # Start of transaction... the clock is ticking...
    [Void]$Socket.Send($NtpData)
    [Void]$Socket.Receive($NtpData)
    $t4 = Get-Date    # End of transaction time
    $Socket.Close()

    $IntPart = [BitConverter]::ToUInt32($NtpData[43..40],0)   # t3
    $FracPart = [BitConverter]::ToUInt32($NtpData[47..44],0)
    $t3ms = $IntPart * 1000 + ($FracPart * 1000 / 0x100000000)

    $IntPart = [BitConverter]::ToUInt32($NtpData[35..32],0)   # t2
    $FracPart = [BitConverter]::ToUInt32($NtpData[39..36],0)
    $t2ms = $IntPart * 1000 + ($FracPart * 1000 / 0x100000000)

    $t1ms = ([TimeZoneInfo]::ConvertTimeToUtc($t1) - $StartOfEpoch).TotalMilliseconds
    $t4ms = ([TimeZoneInfo]::ConvertTimeToUtc($t4) - $StartOfEpoch).TotalMilliseconds

    $Offset = (($t2ms - $t1ms) + ($t3ms-$t4ms))/2

    [String]$NTPDateTime = $StartOfEpoch.AddMilliseconds($t4ms + $Offset).ToLocalTime()

    Try{
        Write-LogEntry ("Synchronizing with NTP server [{0}]." -f $sNTPServer) -Severity 4 -Outhost
        Write-LogEntry ("Attempting to change date and time to: [{0}]..." -f $NTPDateTime) -Severity 4 -Outhost
        Set-Date $NTPDateTime -ErrorAction Stop | Out-Null
        Write-LogEntry ("Successfully updated date and time!") -Severity 4 -Outhost
    }
    Catch{
        Write-LogEntry ("Unable to set date and time: {0}" -f $_.Exception.Message) -Severity 2 -Outhost
    }
}


Function Get-GeographicData {
    [CmdletBinding()] #This provides the function with the -Verbose and -Debug parameters
    param(
        [string]$IpStackAPIKey,
        [string]$BingMapsAPIKey
    )

    #determine if device is intune managed
    #if it is, it will parse log to determine if this script it being ran by intune
    #and remove API sensitive data
    $intuneManagementExtensionLogPath = "$env:ProgramData\Microsoft\IntuneManagementExtension\Logs\IntuneManagementExtension.log"
    If(Test-Path $intuneManagementExtensionLogPath -ErrorAction SilentlyContinue){
        $IntuneManaged = $true
    }else {
        $IntuneManaged = $false
    }

    #attempt connecting online if both keys exist
    If($PSBoundParameters.ContainsKey('IpStackAPIKey') -and $PSBoundParameters.ContainsKey('BingMapsAPIKey'))
    {
        Write-LogEntry "Checking GEO Coordinates by IP for time zone..." -Severity 4 -Outhost
        #Write-Verbose "IPStack API: $IpStackAPIKey"
        #Write-Verbose "Bing Maps API: $BingMapsAPIKey"

        #grab public IP and its geo location
        try {
            $IPStackURI = "http://api.ipstack.com/check?access_key=$($IpStackAPIKey)"
            If($DebugPreference){
                Write-LogEntry ("Initializing Ipstack REST URI: {0}" -f $IPStackURI) -Severity 4 -Outhost
            }Else{
                Write-LogEntry ("Initializing Ipstack REST URI: {0}" -f ($IPStackURI.replace($IpStackAPIKey,'<sensitive data>') ) ) -Severity 4 -Outhost
            }
            $geoIP = Invoke-RestMethod -Uri $IPStackURI -ErrorAction Stop
        }
        Catch {
            Write-LogEntry ("Error obtaining coordinates or public IP address. {0}" -f $_.Exception.Message) -Severity 2 -Outhost
        }
        Finally{
            If($IntuneManaged){
                Write-LogEntry ("Clearing sensitive data in Intune Management Extension log...") -Severity 4 -Outhost
                # Hide the api keys from logs to prevent manipulation API's
                (Get-Content -Path $intuneManagementExtensionLogPath).replace($IpStackAPIKey,'<sensitive data>') |
                            Set-Content -Path $intuneManagementExtensionLogPath -ErrorAction SilentlyContinue | Out-Null
            }
        }

        #determine geo location's timezone
        try {
            Write-LogEntry ("Discovered [{0}] is located in [{1}] at coordinates [{2},{3}]" -f $geoIP.ip,$geoIP.country_name,$geoIP.latitude,$geoIP.longitude) -Severity 4 -Outhost
            $bingURI = "https://dev.virtualearth.net/REST/v1/timezone/$($geoIP.latitude),$($geoIP.longitude)?key=$($BingMapsAPIKey)"
            If($DebugPreference){
                Write-LogEntry ("Initializing BingMaps REST URI: {0}" -f $bingURI) -Severity 4 -Outhost
            }Else{
                Write-LogEntry ("Initializing Ipstack REST URI: {0}" -f ($bingURI.replace($BingMapsAPIKey,'<sensitive data>') ) ) -Severity 4 -Outhost
            }

            $BingApiResponse = Invoke-RestMethod -Uri $bingURI -ErrorAction Stop
            $GEOTimeZone = $BingApiResponse.resourceSets.resources.timeZone.windowsTimeZoneId
            $GEODateTime = $BingApiResponse.resourceSets.resources.timeZone.ConvertedTime | Select -ExpandProperty localTime
        }
        catch {
            Write-LogEntry ("Error obtaining response from Bing Maps API. {0}" -f $_.Exception.Message) -Severity 2 -Outhost
        }
        Finally{
            If($IntuneManaged){
                Write-LogEntry ("Clearing sensitive data in Intune Management Extension log...") -Severity 4 -Outhost
                # Hide the api keys from logs to prevent manipulation API's
                (Get-Content -Path $intuneManagementExtensionLogPath).replace($BingMapsAPIKey,'<sensitive data>') |
                        Set-Content -Path $intuneManagementExtensionLogPath -ErrorAction SilentlyContinue | Out-Null
            }
        }
    }

    If($GEOTimeZone)
    {
        Write-LogEntry ("Discovered geographic time zone: {0}" -f $GEOTimeZone) -Severity 4 -Outhost
        $SelectedTimeZone = $Global:AllTimeZones | Where id -eq $GEOTimeZone
    }Else
    {
        Write-LogEntry ("No geographic time was provided, using current time zone instead...") -Severity 4 -Outhost
        $SelectedTimeZone = $Global:CurrentTimeZone
        $GEOTimeZone = $SelectedTimeZone.Id
    }

    #build GEO data object
    $GeoData = "" | Select DateTime,Id,DisplayName,StandardName
    $GeoData.DateTime = $GEODateTime
    $GeoData.Id = $GEOTimeZone
    $GeoData.DisplayName = $SelectedTimeZone.DisplayName
    $GeoData.StandardName = $SelectedTimeZone.StandardName

    #return data object
    return $GeoData
}


Function Update-DeviceTimeZone{
    <#TEST VALUES
    $SelectedTZ=$ui_lbxTimeZoneList.SelectedItem
    $DefaultTimeZone=(Get-TimeZone).DisplayName
    #>
    [CmdletBinding()]
    param(
        [string]$SelectedTZ
    )
    #update time zone if different than detected
    $SelectedTimeZoneObj = $Global:AllTimeZones | Where {$_.DisplayName -eq $SelectedTZ}

    If($SelectedTZ -ne $Global:CurrentTimeZone.DisplayName)
    {
        Try{
            Write-LogEntry ("Attempting to change time zone to: {0}..." -f $SelectedTZ) -Severity 4 -Outhost
            Set-TimeZone $SelectedTimeZoneObj -ErrorAction Stop | Out-Null
            Start-Service W32Time | Restart-Service -ErrorAction Stop
            Write-LogEntry ("Completed time zone change!" -f $SelectedTZ) -Severity 4 -Outhost
        }
        Catch{
            #Throw $_.Exception.Message
            Write-LogEntry ("Failed to set device time zone. {0}" -f $_.Exception.Message) -Severity 3 -Outhost
            Exit $_.Exception.HResult
        }
    }Else{
        Write-LogEntry "No change. Skipping time zone update" -Severity 4 -Outhost
    }
}

#===========================================================================
# Actually make the UI work
#===========================================================================
#splat Params. Check if IPstack and Bingmap values DO NOT EXIST; use default timeseletions
If(  ([string]::IsNullOrEmpty($IpStackAPIKey)) -or ([string]::IsNullOrEmpty($BingMapsAPIKey)) ){
    $ui_txtTimeZoneTitle.Text = $ui_txtTimeZoneTitle.Text -replace "@anchor","What time zone are you in?"
    $GeoTZParams = @{
        Verbose=$VerbosePreference
        Debug=$DebugPreference
    }
}
Else{
    $ui_txtTimeZoneTitle.Text = $ui_txtTimeZoneTitle.Text -replace "@anchor","Is this the time zone your in?"
    $GeoTZParams = @{
        ipStackAPIKey=$IpStackAPIKey
        bingMapsAPIKey=$BingMapsAPIKey
        Verbose=$VerbosePreference
        Debug=$DebugPreference
    }
}

#determine if Selector control will update status
If(!(Test-WinPE) -and ($NoControl -eq $false))
{
    $UIControlParam = @{
        UIObject=$TZSelectUI
        UpdateStatusKeyHive=$RegHive
        Verbose=$VerbosePreference
        Debug=$DebugPreference
    }
}Else{
    $UIControlParam = @{
        UIObject=$TZSelectUI
        Verbose=$VerbosePreference
        Debug=$DebugPreference
    }
}

#Get all timezones and load it to combo box
$Global:AllTimeZones.DisplayName | ForEach-object {$ui_lbxTimeZoneList.Items.Add($_)} | Out-Null

#grab Geo Timezone
<#TEST Timezones
$TargetGeoTzObj = (Get-TimeZone -listAvailable)[9] #<---(UTC-08:00) Pacific Time (US & Canada)
$TargetGeoTzObj = (Get-TimeZone -listAvailable)[12] #<---(UTC-07:00) Mountain Time (US & Canada)
$TargetGeoTzObj = (Get-TimeZone -listAvailable)[15] #<---(UTC-06:00) Central Time (US & Canada)
$TargetGeoTzObj = (Get-TimeZone -listAvailable)[21] #<---(UTC-05:00) Eastern Time (US & Canada)
$TargetGeoTzObj = (Get-TimeZone)
#>
$TargetGeoTzObj = Get-GeographicData @GeoTZParams

#select current time zone
Write-LogEntry ("The selected time zone is: {0}" -f $TargetGeoTzObj.DisplayName) -Severity 4 -Outhost
$ui_lbxTimeZoneList.SelectedItem = $TargetGeoTzObj.DisplayName

#scrolls list to current selected item
#+3 below to center selected item on screen
$ui_lbxTimeZoneList.ScrollIntoView($ui_lbxTimeZoneList.Items[$ui_lbxTimeZoneList.SelectedIndex+3])

#when button is clicked changer time
$ui_btnTZSelect.Add_Click({
    #Set time zone
    #Set-TimeZone $ui_lbxTimeZoneList.SelectedItem
    Update-DeviceTimeZone -SelectedTZ $ui_lbxTimeZoneList.SelectedItem
    #build registry key for time selector
    If($null -ne $UIControlParam.UpdateStatusKeyHive){Set-StatusKey -Hive $RegHive -Name TimeZoneSelected -Value $ui_lbxTimeZoneList.SelectedItem}

	#update the time and date
    If($SyncNTP)
    {
        Set-NTPDateTime -sNTPServer $Global:NTPServer -Verbose:$VerbosePreference
        If($null -ne $UIControlParam.UpdateStatusKeyHive){Set-StatusKey -Hive $RegHive -Name SyncedToNTP -Value $Global:NTPServer}
    }
    Else{
        Write-LogEntry ("No NTP server specified. Skipping date and time update.") -Severity 4 -Outhost
    }
    #close the UI
    Stop-TimeSelectorUI @UIControlParam
	#If(!$isISE){Stop-Process $pid}

})


#===========================================================================
# Main - Call the form depending on logic
#===========================================================================
<# LOGIC for UI

UI will show if:
 - 'ForceInteraction' parameter set to True
 - 'UI has not detected itself running before or it has failed or not completed

UI will NOT show if:
 - 'ForceInteraction' parameter set to False AND 'NoUI' parameter is set to True
 - Time is not different from detected Geographic time (this will only work if IpStack/Bing API are included)
 - If script detected its still running OR last status is set to 'running'

UI will make changes if:
- If UI is displayed and change is selected
- NoUI is set to True

#>

# found that if script is called by Intune, the script may be running multiple times if the ESP screen process takes a while
# Only allow the script to run once if it is already being displayed
If($ForceInteraction){
    #run form all the time
    Write-LogEntry ("'ForceInteraction' parameter is enabled; UI will be displayed") -Severity 4 -Outhost
    Start-TimeSelectorUI @UIControlParam
}
#if noUI is set; attempt to set the timezone and time without UI interaction
ElseIf($NoUI)
{
    Write-LogEntry ("'NoUI' parameter is enabled; UI will NOT be displayed") -Severity 4 -Outhost
    #update the time zone
    Update-DeviceTimeZone -SelectedTZ $TargetGeoTzObj.DisplayName

    #log changes to registry
    If($null -ne $UIControlParam.UpdateStatusKeyHive){Set-StatusKey -Hive $RegHive -Name TimeZoneSelected -Value $ui_lbxTimeZoneList.SelectedItem}

    #update the time and date
    If($SyncNTP){
        Set-NTPDateTime -sNTPServer $Global:NTPServer -Verbose:$VerbosePreference
        If($null -ne $UIControlParam.UpdateStatusKeyHive){Set-StatusKey -Hive $RegHive -Name SyncedToNTP -Value $Global:NTPServer}
    }Else{
        Write-LogEntry ("No NTP server specified. Skipping date and time update.") -Severity 4 -Outhost
    }

}
Elseif($NoControl){
    Write-LogEntry ("'NoControl' is enabled; UI will be displayed without monitoring registry.") -Severity 4 -Outhost
    Start-TimeSelectorUI @UIControlParam
}
ElseIf( Get-Process | Where {$_.MainWindowTitle -eq "Time Zone Selection"} ){
    #do nothing
    Write-LogEntry "Detected that UI process is still running. UI will not be displayed." -Severity 4 -Outhost
}
ElseIf($RunOnce){
    $UiStatus = Get-ItemPropertyValue "$($RegHive):\Software\PowerShellCrack\TimeZoneSelector" -Name Status -ErrorAction SilentlyContinue
    switch($UiStatus){
        'Running' {$StatusMsg = "Script must have crashed because process detection was not found and status is running. UI will be displayed"; $displayUI = $true}
        'Failed' {$StatusMsg = "Last attempt failed; UI will be displayed."; $displayUI = $true}
        'Completed' {$StatusMsg = "Selector has already ran once. Try '-ForceInteraction:`$true' param to force the UI."; $displayUI = $false}
        $null {$StatusMsg = "First time running script; UI will be displayed."; $displayUI = $true}
        default {$StatusMsg = "Unknown status; UI will be displayed."; $displayUI = $true}
    }

    #check if registry key exists to determine if form needs to be displayed
    If($displayUI){
        Write-LogEntry $StatusMsg -Severity 4 -Outhost
        Start-TimeSelectorUI @UIControlParam
    }Else{
        #do nothing
        #Stop-TimeSelectorUI -UIObject $TZSelectUI -UpdateStatusKeyHive $RegHive -CustomStatus "Completed"
        Write-LogEntry $StatusMsg -Severity 4 -Outhost
    }
}
ElseIf($TargetGeoTzObj.DisplayName -ne $Global:CurrentTimeZone.DisplayName){
    #Only run if time compared differs
    Write-LogEntry ("Current time is different than Geo time scenario; UI will be displayed") -Severity 4 -Outhost
    Start-TimeSelectorUI @UIControlParam
}
Else{
    Write-LogEntry ("All scenarios are false; UI will be displayed") -Severity 4 -Outhost
    Start-TimeSelectorUI @UIControlParam
}



#if running in a Task sequence output the timezone to standard TS variables
#https://docs.microsoft.com/en-us/mem/configmgr/osd/understand/task-sequence-variables#OSDTimeZone-output
If( (Test-SMSTSENV) -and ($ui_lbxTimeZoneList.SelectedItem) )
{
    Write-LogEntry ("Task Sequence detected, settings output variables: ") -Severity 4 -Outhost
    Write-LogEntry ("OSDMigrateTimeZone: {0}" -f $True.ToString()) -Severity 4 -Outhost
    Write-LogEntry ("OSDTimeZone: {0}" -f $TargetGeoTzObj.StandardName) -Severity 4 -Outhost
    #$tsenv.Value("TimeZone") = (Get-TimeZoneIndex -TimeZone $ui_lbxTimeZoneList.SelectedItem #<--- TODO Need index function created
    $tsenv.Value("OSDMigrateTimeZone") = $true
    $tsenv.Value("OSDTimeZone") = $TargetGeoTzObj.StandardName
}