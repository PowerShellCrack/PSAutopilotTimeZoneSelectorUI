
<#
    .SYNOPSIS
        Prompts user to set time zone

    .DESCRIPTION
		Prompts user to set time zone using windows presentation framework
        Can be used in:
            - SCCM Tasksequences (User interface allowed)
            - SCCM Software Delivery (User interface allowed)
            - Intune Autopilot

    .NOTES
        Author		: Dick Tracy <richard.tracy@hotmail.com>
	    Source		: https://github.com/PowerShellCrack/AutopilotTimeZoneSelectorUI
        Version		: 2.0.0
        README      : Review README.md for more details and configurations
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
        PS> .\TimeZoneUI.ps1 -IpStackAPIKey "4bd1443445dfhrrt9dvefr45341" -BingMapsAPIKey "Bh53uNUOwg71czosmd73hKfdHf465ddfhrtpiohvknlkewufjf4-d" -Verbose

        RESULT: Uses IP GEO location for the pre-selection

    .EXAMPLE
        PS> .\TimeZoneUI.ps1 -ForceInteraction:$true -verbose

        RESULT:  This will ALWAYS display the time selection screen; if IPStack and BingMapsAPI included the IP GEO location timezone will be preselected. Verbose output will be displayed

    .EXAMPLE
        PS> .\TimeZoneUI.ps1 -IpStackAPIKey "4bd1443445dfhrrt9dvefr45341" -BingMapsAPIKey "Bh53uNUOwg71czosmd73hKfdHf465ddfhrtpiohvknlkewufjf4-d" -NoUI:$true -SyncNTP "time-a-g.nist.gov"

        RESULT: This will set the time automatically using the IP GEO location without prompting user. If API not provided, timezone or time will not change the current settings

    .EXAMPLE
        PS> .\TimeZoneUI.ps1 -UserDriven:$false

        RESULT: Writes a registry key in System (HKEY_LOCAL_MACHINE) hive to determine run status

    .EXAMPLE
        PS> .\TimeZoneUI.ps1 -RunOnce:$true

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
        Title="Time Zone Selection">
    <Window.Resources>
        <ResourceDictionary>

            <Style TargetType="{x:Type Window}">
                <Setter Property="FontFamily" Value="Segoe UI" />
                <Setter Property="FontWeight" Value="Light" />
                <Setter Property="Background" Value="#FF1D3245" />
                <Setter Property="Foreground" Value="#FFE8EDF9" />
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
            <Style TargetType="{x:Type Button}">
                <Setter Property="Background" Value="#FF1D3245" />
                <Setter Property="Foreground" Value="#FFE8EDF9" />
                <Setter Property="FontSize" Value="15" />
                <Setter Property="SnapsToDevicePixels" Value="True" />

                <Setter Property="Template">
                    <Setter.Value>
                        <ControlTemplate TargetType="Button" >

                            <Border Name="border"
                                BorderThickness="1"
                                Padding="4,2"
                                BorderBrush="#FF1D3245"
                                CornerRadius="2"
                                Background="#00A4EF">
                                <ContentPresenter HorizontalAlignment="Center"
                                                VerticalAlignment="Center"
                                                TextBlock.TextAlignment="Center"
                                                />
                            </Border>

                            <ControlTemplate.Triggers>
                                <Trigger Property="IsMouseOver" Value="True">
                                    <Setter TargetName="border" Property="BorderBrush" Value="#FFE8EDF9" />
                                </Trigger>

                                <Trigger Property="IsPressed" Value="True">
                                    <Setter TargetName="border" Property="BorderBrush" Value="#FF1D3245" />
                                    <Setter Property="Button.Foreground" Value="#FF1D3245" />
                                    <Setter Property="Effect">
                                        <Setter.Value>
                                            <DropShadowEffect ShadowDepth="0" Color="#FF1D3245" Opacity="1" BlurRadius="10"/>
                                        </Setter.Value>
                                    </Setter>
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
                            <Border Name="Border" BorderThickness="1" CornerRadius="2">
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
                            <Border Name="ItemBorder" Padding="8" Margin="1" Background="#FF1D3245">
                                <ContentPresenter />
                            </Border>
                            <ControlTemplate.Triggers>
                                <Trigger Property="IsSelected" Value="True">
                                    <Setter TargetName="ItemBorder" Property="Background" Value="#00A4EF" />
                                    <Setter Property="Foreground" Value="#FF1D3245" />
                                </Trigger>
                                <MultiTrigger>
                                    <MultiTrigger.Conditions>
                                        <Condition Property="IsMouseOver" Value="True" />
                                        <Condition Property="IsSelected" Value="False" />
                                    </MultiTrigger.Conditions>
                                    <Setter TargetName="ItemBorder" Property="Background" Value="#00A4EF" />
                                </MultiTrigger>
                            </ControlTemplate.Triggers>
                        </ControlTemplate>
                    </Setter.Value>
                </Setter>
            </Style>
        </ResourceDictionary>
    </Window.Resources>

    <Grid Background="#FF1D3245" HorizontalAlignment="Center" VerticalAlignment="Center" Height="600">

        <TextBlock x:Name="txtTimeZoneTitle" HorizontalAlignment="Center" Text="@anchor" VerticalAlignment="Top" FontSize="48"/>
        <ListBox x:Name="lbxTimeZoneList" HorizontalAlignment="Center" VerticalAlignment="Top" Background="#FF1D3245" Foreground="#FFE8EDF9" FontSize="18" Width="700" Height="400" Margin="0,80,0,0" ScrollViewer.VerticalScrollBarVisibility="Visible" SelectionMode="Single"/>

        <Button x:Name="btnTZSelect" Content="Select Time Zone" Height="65" Width="200" HorizontalAlignment="Center" VerticalAlignment="Bottom" FontSize="18" Padding="10"/>

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

# Only set keys if not in PE AND NoControl is not enabled
If(!(Test-WinPE) -and ($NoControl -eq $false))
{
    #Set registry hive for user or local machine
    If($UserDriven -eq $false){$RegHive = 'HKLM:'}Else{$RegHive = 'HKCU:'}
    $RegPath = "SOFTWARE\PowerShellCrack\TimeZoneSelector"
    # Build registry key for status and selection
    #if unable to create key, deployment or permission may need to change
    Try{
        If(-not(Test-Path "$RegHive\$RegPath") ){
            New-Item -Path "$RegHive\SOFTWARE" -Name "PowerShellCrack" -ErrorAction SilentlyContinue | Out-Null
            New-Item -Path "$RegHive\SOFTWARE\PowerShellCrack" -Name "TimeZoneSelector" -ErrorAction Stop | Out-Null
        }
    }
    Catch{
        Write-LogEntry ("Unable to set registry key [{0}\{1}] with value [{2}]. {3}" -f "$RegHive\$RegPath", "TimeZoneSelector", $TargetTimeZone.id,$_.Exception.Message) -Severity 3 -Outhost
        Exit $_.Exception.HResult
    }
}
#====================
#Form Functions
#====================


Function Start-TimeSelectorUI{
    <#TEST VALUES
    $UIObject=$TZSelectUI
    $UpdateStatusKey="$RegHive\$RegPath"
    $UpdateStatusKey="HKLM:\SOFTWARE\PowerShellCrack\TimeZoneSelector"
    #>
    [CmdletBinding()]
    param(
        $UIObject,
        [string]$UpdateStatusKey
    )
    If($PSBoundParameters.ContainsKey('UpdateStatusKey')){Set-ItemProperty -Path $UpdateStatusKey -Name Status -Value "Running" -Force -ErrorAction SilentlyContinue}

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
        If($PSBoundParameters.ContainsKey('UpdateStatusKey')){Set-ItemProperty -Path $UpdateStatusKey -Name Status -Value 'Failed' -Force -ErrorAction SilentlyContinue}
        Write-LogEntry ("Unable to load Windows Presentation UI. {0}" -f $_.Exception.Message) -Severity 3 -Outhost
        Exit $_.Exception.HResult
    }
}

function Stop-TimeSelectorUI{
    <#TEST VALUES
    $UIObject=$TZSelectUI
    $UpdateStatusKey="$RegHive\$RegPath"
    $UpdateStatusKey="HKLM:\SOFTWARE\PowerShellCrack\TimeZoneSelector"
    #>
    [CmdletBinding()]
    param(
        $UIObject,
        [string]$UpdateStatusKey,
        [string]$CustomStatus
    )

    If($CustomStatus){$status = $CustomStatus}
    Else{$status = 'Completed'}

    Try{
        If($PSBoundParameters.ContainsKey('UpdateStatusKey')){Set-ItemProperty -Path $UpdateStatusKey -Name Status -Value $status -Force -ErrorAction SilentlyContinue}
        #$UIObject.Close() | Out-Null
        $UIObject.Close()
    }
    Catch{
        If($PSBoundParameters.ContainsKey('UpdateStatusKey')){Set-ItemProperty -Path $UpdateStatusKey -Name Status -Value 'Failed' -Force -ErrorAction SilentlyContinue}
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
# Actually make the objects work
#===========================================================================

#find a time zone to select

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
        UpdateStatusKey="$RegHive\$RegPath"
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
    If($null -ne $UIControlParam.UpdateStatusKey){Set-ItemProperty -Path "$RegHive\$RegPath" -Name TimeZoneSelected -Value $ui_lbxTimeZoneList.SelectedItem -Force -ErrorAction SilentlyContinue}

	#update the time and date
    If($SyncNTP)
    {
        Set-NTPDateTime -sNTPServer $Global:NTPServer -Verbose:$VerbosePreference
        If($null -ne $UIControlParam.UpdateStatusKey){Set-ItemProperty -Path "$RegHive\$RegPath" -Name NTPTimeSynced -Value $Global:NTPServer -Force -ErrorAction SilentlyContinue}
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
    If($null -ne $UIControlParam.UpdateStatusKey){Set-ItemProperty -Path "$RegHive\$RegPath" -Name TimeZoneSelected -Value $ui_lbxTimeZoneList.SelectedItem -Force -ErrorAction SilentlyContinue}

    #update the time and date
    If($SyncNTP){
        Set-NTPDateTime -sNTPServer $Global:NTPServer -Verbose:$VerbosePreference
        If($null -ne $UIControlParam.UpdateStatusKey){Set-ItemProperty -Path "$RegHive\$RegPath" -Name NTPTimeSynced -Value $Global:NTPServer -Force -ErrorAction SilentlyContinue}
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
    $UiStatus = (Get-ItemProperty "$RegHive\$RegPath" -Name Status -ErrorAction SilentlyContinue).Status
    switch($UiStatus){
        #"Running" {$StatusMsg = "Script status shows 'Running'.  UI will not be displayed"; $displayUI = $false}
        "Failed" {$StatusMsg = "Last attempt failed; UI will be displayed."; $displayUI = $true}
        "Completed" {$StatusMsg = "Selector has already ran once. Try '-ForceInteraction:`$true' param to force the UI."; $displayUI = $false}
        $null {$StatusMsg = "First time running script; UI will be displayed."; $displayUI = $true}
        default {$StatusMsg = "Unknown status; UI will be displayed."; $displayUI = $true}
    }
    #check if registry key exists to determine if form needs to be displayed\
    If($displayUI){
        Write-LogEntry $StatusMsg -Severity 4 -Outhost
        Start-TimeSelectorUI @UIControlParam
    }Else{
        #do nothing
        #Stop-TimeSelectorUI -UIObject $TZSelectUI -UpdateStatusKey "$RegHive\$RegPath"-CustomStatus "Completed"
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