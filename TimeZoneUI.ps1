
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
        Launches in full screen using WPF

    .LINK
        https://matthewjwhite.co.uk/2019/04/18/intune-automatically-set-timezone-on-new-device-build/
        https://ipstack.com
        https://azuremarketplace.microsoft.com/en-us/marketplace/apps/bingmaps.mapapis

    .PARAMETER IpStackAPIKey
        Used to get geoCoordinates of the public IP. get the API key from https://ipstack.com

    .PARAMETER BingMapsAPIKeyy
        Used to get the Windows TimeZone value of the location coordinates. get the API key from https://azuremarketplace.microsoft.com/en-us/marketplace/apps/bingmaps.mapapis

    .PARAMETER UserDriven
        deploy to user sets either HKCU key or HKLM key
        Set to true if the deployment is for autopilot
        NOTE: Permission required for HKLM

    .PARAMETER OnlyRunOnce
        Specify that this script will only launch the form one time.

    .PARAMETER ForceTimeSelection
        Disabled and with Bing API --> Current timezone and geo timezone will be compared; if different, form will be displayed
        Enabled --> the selection will always show

    .PARAMETER AutoTimeSelection
        Enabled with Bing API --> No prompt for user, time will update on it own
        Enabled without Bing API --> User will be prompted at least once
        Ignored if ForceTimeSelection is enabled

    .PARAMETER UpdateTime
        Used only with IPstack and Bing API
        Set local time and date (NOT TIMEZONE) based on GEO location
        Requires administrative permissions

    .EXAMPLE
        PS> .\TimeZoneUI.ps1 -IpStackAPIKey = "4bd1443445dfhrrt9dvefr45341" -BingMapsAPIKey = "Bh53uNUOwg71czosmd73hKfdHf465ddfhrtpiohvknlkewufjf4-d" -Verbose

        Uses IP GEO location for the pre-selection

    .EXAMPLE
        PS> .\TimeZoneUI.ps1 -ForceTimeSelection

        This will always display the time selection screen; if IPStack and BingMapsAPI included the IP GEO location timezone will be preselected

    .EXAMPLE
        PS> .\TimeZoneUI.ps1 -IpStackAPIKey = "4bd1443445dfhrrt9dvefr45341" -BingMapsAPIKey = "Bh53uNUOwg71czosmd73hKfdHf465ddfhrtpiohvknlkewufjf4-d" -AutoTimeSelection -UpdateTime

        This will set the time automatically using the IP GEO location without prompting user. If API not provided, timezone or time will not change the current settings

    .EXAMPLE
        PS> .\TimeZoneUI.ps1 -UserDriven $false

        Writes a registry key in HKLM hive to determine run status

    .EXAMPLE
        PS> .\TimeZoneUI.ps1 -OnlyRunOnce $true

        Mainly for Autopilot powershell scripts; this allows the screen to display one time after ESP is completed.
#>

#===========================================================================
# CONTROL VARIABLES
#===========================================================================

[CmdletBinding()]
param(
    [string]$IpStackAPIKey = "",

    [string]$BingMapsAPIKey = "" ,

    [boolean]$UserDriven = $true,

    [boolean]$OnlyRunOnce = $true,

    [switch]$ForceTimeSelection,

    [switch]$AutoTimeSelection,

    [switch]$UpdateTime
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

##*=============================================
##* VARIABLE DECLARATION
##*=============================================
#region VARIABLES: Building paths & values
[string]$scriptPath = Get-ScriptPath

#Grab all times zones plus current timezones
$Global:AllTimeZones = Get-TimeZone -ListAvailable
$Global:CurrentTimeZone = Get-TimeZone
#===========================================================================
# XAML LANGUAGE
#===========================================================================
$inputXML = @"
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

    <Grid x:Name="background" HorizontalAlignment="Center" VerticalAlignment="Center" Height="600">

        <TextBlock x:Name="targetTZ_label" HorizontalAlignment="Center" Text="@anchor" VerticalAlignment="Top" FontSize="48"/>
        <ListBox x:Name="targetTZ_listBox" HorizontalAlignment="Center" VerticalAlignment="Top" Background="#FF1D3245" Foreground="#FFE8EDF9" FontSize="18" Width="700" Height="400" Margin="0,80,0,0" ScrollViewer.VerticalScrollBarVisibility="Visible" SelectionMode="Single"/>
        <Grid x:Name="msg" Width="700" Height="100" Margin="0,360,0,0" HorizontalAlignment="Center">
            <Grid.ColumnDefinitions>
                <ColumnDefinition Width="1*" />
                <ColumnDefinition Width="1*" />
            </Grid.ColumnDefinitions>
            <!--
            <TextBlock x:Name="DefaultTZMsg" Grid.Column="0" Text="If a time zone is not selected, time will be set to: " HorizontalAlignment="Right" VerticalAlignment="Bottom" FontSize="16" Foreground="#00A4EF"/>
            <TextBlock x:Name="CurrentTZ" Grid.Column="1" Text="@anchor" HorizontalAlignment="Left" VerticalAlignment="Bottom" FontSize="16" Foreground="yellow"/>
            -->
        </Grid>
        <Button x:Name="ChangeTZButton" Content="Select Time Zone" Height="65" Width="200" HorizontalAlignment="Center" VerticalAlignment="Bottom" FontSize="18" Padding="10"/>

    </Grid>
</Window>
"@

#replace some default attributes to support powershell
$inputXML = $inputXML -replace 'mc:Ignorable="d"','' -replace "x:N",'N'  -replace '^<Win.*', '<Window'

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

[xml]$XAML = $inputXML
#Read XAML
$reader=(New-Object System.Xml.XmlNodeReader $xaml)
try{$TZSelectUI=[Windows.Markup.XamlReader]::Load( $reader )}
catch{Write-Host "Unable to load Windows.Markup.XamlReader. Double-check syntax and ensure .net is installed."}

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
    Throw ("Unable to configure registry key [{0}\{1}]. {3}" -f "$RegHive\$RegPath", 'TimeZoneSelected ',$TargetTimeZone.id,$_.Exception.Message)
    Exit -1
}

#====================
#Form Functions
#====================


function Start-TimeSelectorUI{
    <#TEST VALUES
    $UIObject=$TZSelectUI
    $UpdateStatusKey="$RegHive\$RegPath"
    $UpdateStatusKey="HKLM:\SOFTWARE\PowerShellCrack\TimeZoneSelector"
    #>
    param(
        [CmdletBinding()]
        $UIObject,
        [string]$UpdateStatusKey
    )
    If($UpdateStatusKey){Set-ItemProperty -Path $UpdateStatusKey -Name Status -Value "Running" -Force -ErrorAction SilentlyContinue}

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
        If($UpdateStatusKey){Set-ItemProperty -Path $UpdateStatusKey -Name Status -Value 'Failed' -Force -ErrorAction SilentlyContinue}
        Throw $_.Exception.Message
    }
}

function Stop-TimeSelectorUI{
    <#TEST VALUES
    $UIObject=$TZSelectUI
    $UpdateStatusKey="$RegHive\$RegPath"
    $UpdateStatusKey="HKLM:\SOFTWARE\PowerShellCrack\TimeZoneSelector"
    #>
    param(
        [CmdletBinding()]
        $UIObject,
        [string]$UpdateStatusKey,
        [string]$CustomStatus
    )

    If($CustomStatus){$status = $CustomStatus}
    Else{$status = 'Completed'}

    If($UpdateStatusKey){Set-ItemProperty -Path $UpdateStatusKey -Name Status -Value $status -Force -ErrorAction SilentlyContinue}

    Try{
        #$UIObject.Close() | Out-Null
        $UIObject.Close()
    }
    Catch{
        If($UpdateStatusKey){Set-ItemProperty -Path $UpdateStatusKey -Name Status -Value 'Failed' -Force -ErrorAction SilentlyContinue}
        $_.Exception.Message
    }
}


function Set-NTPDateTime
{
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

    set-date $NTPDateTime
}


function Get-GEOTimeZone {
    param(
        [CmdletBinding()]
        [string]$IpStackAPIKey,
        [string]$BingMapsAPIKey,
        [boolean]$ChangeTimeDate
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
        Write-Verbose "Checking GEO Coordinates by IP for timezone..."
        Write-Verbose "IPStack API: $IpStackAPIKey"
        Write-Verbose "Bing Maps API: $BingMapsAPIKey"

        #grab public IP and its geo location
        try {
            $IPStackURI = "http://api.ipstack.com/check?access_key=$($IpStackAPIKey)"
            $geoIP = Invoke-RestMethod -Uri $IPStackURI -ErrorAction Stop
        }
        Catch {
            Write-Verbose "Error obtaining coordinates or public IP address"
        }
        Finally{
            If($IntuneManaged){
                # Hide the api keys from logs to prevent manipulation API's
                (Get-Content -Path $intuneManagementExtensionLogPath).replace($IpStackAPIKey,'<sensitive data>') |
                            Set-Content -Path $intuneManagementExtensionLogPath -ErrorAction SilentlyContinue | Out-Null
            }
        }

        #determine geo location's timezone
        try {
            Write-Verbose "Detected that $($geoIP.ip) is located in $($geoIP.country_name) at $($geoIP.latitude),$($geoIP.longitude)"
            $bingURI = "https://dev.virtualearth.net/REST/v1/timezone/$($geoIP.latitude),$($geoIP.longitude)?key=$($BingMapsAPIKey)"
            $BingApiResponse = Invoke-RestMethod -Uri $bingURI -ErrorAction Stop
        }
        catch {
            Write-Verbose ("Error obtaining response from Bing Maps API. {0}" -f $_.exception.message)
        }
        Finally{
            If($IntuneManaged){
                # Hide the api keys from logs to prevent manipulation API's
                (Get-Content -Path $intuneManagementExtensionLogPath).replace($BingMapsAPIKey,'<sensitive data>') |
                        Set-Content -Path $intuneManagementExtensionLogPath -ErrorAction SilentlyContinue | Out-Null
            }
            $correctTimeZone = $BingApiResponse.resourceSets.resources.timeZone.windowsTimeZoneId
            Write-Verbose "Detected Correct time zone as '$($correctTimeZone)'"
            If($correctTimeZone){$SelectedTimeZone = [string](Get-TimeZone -id $correctTimeZone).DisplayName}

            If($ChangeTimeDate){
                $geoTimeDate = $BingApiResponse.resourceSets.resources.timeZone.ConvertedTime | Select -ExpandProperty localTime
                try {
                    Write-Verbose ("Attempting to set local time to: {0}" -f (Get-Date $geoTimeDate))
                    If($geoTimeDate)
                    {
                        Set-Date $geoTimeDate -ErrorAction Stop
                    }
                    Else{
                        Set-NTPDateTime -sNTPServer 'pool.ntp.org'
                    }
                }
                catch {
                    Write-Verbose ("Error setting time and date from Bing Maps API: {0}" -f $_.Exception.Message)
                }
            }
        }
    }
    Else
    {
        Write-Verbose ("Offline Time Zone detection will run...")
    }

    #confirm if time zone value exists, if not default to current time
    If(!$SelectedTimeZone){
        $SelectedTimeZone = $Global:CurrentTimeZone
    }

    #return selected timezone object
    return ($Global:AllTimeZones | Where {$_.Displayname -eq $SelectedTimeZone})
}

Function Update-DeviceTimeZone{
    <#TEST VALUES
    $SelectedInput=$ui_targetTZ_listBox.SelectedItem
    $DefaultTimeZone=(Get-TimeZone).DisplayName
    #>
    param(
        [string]$SelectedInput,
        [string]$DefaultTimeZone
    )
    #update time zone if different than detected
    $SelectedTimeZoneObj = $Global:AllTimeZones | Where {$_.DisplayName -eq $SelectedInput}
    If($SelectedInput -ne $DefaultTimeZone)
    {
        Try{
            Write-Verbose ("Attempting to chang Time Zone to: {0}..." -f $SelectedInput)
            Set-TimeZone $SelectedTimeZoneObj -ErrorAction SilentlyContinue | Out-Null
            Start-Service W32Time | Restart-Service -ErrorAction SilentlyContinue
            Write-Verbose ("Completed Time Zone change" -f $SelectedInput)
        }
        Catch{
            Throw $_.Exception.Message
        }
    }Else{
        Write-Verbose ("Same Time Zone has been selected: {0}" -f $SelectedInput)
    }
}

#===========================================================================
# Actually make the objects work
#===========================================================================

#find a time zone to select

#splat Params. Check if IPstack and Bingmap values DO NOT EXIST; use default timeseletions
If(  ([string]::IsNullOrEmpty($IpStackAPIKey)) -or ([string]::IsNullOrEmpty($BingMapsAPIKey)) ){
    $ui_targetTZ_label.Text = $ui_targetTZ_label.Text -replace "@anchor","What time zone are you in?"
    $params = @{
        Verbose=$VerbosePreference
    }
}
Else{
    $ui_targetTZ_label.Text = $ui_targetTZ_label.Text -replace "@anchor","Is this the time zone your in?"
    $params = @{
        ipStackAPIKey=$IpStackAPIKey
        bingMapsAPIKey=$BingMapsAPIKey
        Verbose=$VerbosePreference
    }
}

#if set, script will attempt to change time and sat without user intervention
If($PSBoundParameters.ContainsKey('UpdateTime')){
    $params += @{
        ChangeTimeDate=$true
    }
}
Else{
    $params += @{
        ChangeTimeDate=$false
    }
}



#Get all timezones and load it to combo box
$AllTimeZones.DisplayName | ForEach-object {$ui_targetTZ_listBox.Items.Add($_)} | Out-Null

#grab Geo Timezone
<#TEST Timezones
$TargetGeoTzObj = (Get-TimeZone -listAvailable)[9] #<---(UTC-08:00) Pacific Time (US & Canada)
$TargetGeoTzObj = (Get-TimeZone -listAvailable)[12] #<---(UTC-07:00) Mountain Time (US & Canada)
$TargetGeoTzObj = (Get-TimeZone -listAvailable)[15] #<---(UTC-06:00) Central Time (US & Canada)
$TargetGeoTzObj = (Get-TimeZone -listAvailable)[21] #<---(UTC-05:00) Eastern Time (US & Canada)
$TargetGeoTzObj = (Get-TimeZone)
#>
$TargetGeoTzObj = Get-GEOTimeZone @params
Write-Verbose ("Detected Time Zone is: {0}" -f $TargetGeoTzObj.id)

#select current time zone
$ui_targetTZ_listBox.SelectedItem = $TargetGeoTzObj.DisplayName

#scrolls list to current selected item
#+3 below to center selected item on screen
$ui_targetTZ_listBox.ScrollIntoView($ui_targetTZ_listBox.Items[$ui_targetTZ_listBox.SelectedIndex+3])

#if autoselection is enabled, attempt setting the time zone
If($AutoTimeSelection)
{
    Write-Verbose "Auto Selection parameter used"
    Write-Verbose ("Attempting to auto set Time Zone to: {0}..." -f $TargetGeoTzObj.id)

    #update the time zone
    Update-DeviceTimeZone -SelectedInput $ui_targetTZ_listBox.SelectedItem -DefaultTimeZone ($Global:CurrentTimeZone).DisplayName

    #update the time and date
    If($PSBoundParameters.ContainsKey('UpdateTime') ){Set-NTPDateTime -sNTPServer 'pool.ntp.org'}

    #log changes to registry
    Set-ItemProperty -Path "$RegHive\$RegPath" -Name TimeZoneSelected -Value $ui_targetTZ_listBox.SelectedItem -Force -ErrorAction SilentlyContinue
}

#compare the GEO Targeted timezone verses the current timezone
#this is check during run time
If($TargetGeoTzObj.id -eq ((Get-TimeZone).Id) ){
    $TimeComparisonDiffers = $false
}
Else{
    $TimeComparisonDiffers = $true
}

#when button is clicked changer time
$ui_ChangeTZButton.Add_Click({
    #Set time zone
    #Set-TimeZone $ui_targetTZ_listBox.SelectedItem
    Update-DeviceTimeZone -SelectedInput $ui_targetTZ_listBox.SelectedItem -DefaultTimeZone $TargetGeoTzObj.DisplayName
    #build registry key for time selector
    Set-ItemProperty -Path "$RegHive\$RegPath" -Name TimeZoneSelected -Value $ui_targetTZ_listBox.SelectedItem -Force -ErrorAction SilentlyContinue
    #close the UI
    Stop-TimeSelectorUI -UIObject $TZSelectUI -UpdateStatusKey "$RegHive\$RegPath"
	#If(!$isISE){Stop-Process $pid}
})

#===========================================================================
# Main - Call the form depending on scenario
#===========================================================================

# found that if script is called by Intune, the script may be running multiple times if the ESP screen process takes a while
# Only allow the script to run once if it is already being displayed
If($ForceTimeSelection){
    #run form all the time
    Write-Verbose ("'ForceTimeSelection' parameter called; selector will be displayed")
    Start-TimeSelectorUI -UIObject $TZSelectUI -UpdateStatusKey "$RegHive\$RegPath"
}
ElseIf((Get-ItemProperty "$RegHive\$RegPath" -Name Status -ErrorAction SilentlyContinue).Status -eq "Running"){
    #do nothing
    Write-Verbose "Detected that TimeSelector UI is running. Exiting"
}
ElseIf($TimeComparisonDiffers -eq $true){
    #Only run if time compared differs
    Write-Verbose ("Current time is different than Geo time scenario; selector will be displayed")
    Start-TimeSelectorUI -UIObject $TZSelectUI -UpdateStatusKey "$RegHive\$RegPath"
}
ElseIf($OnlyRunOnce){
    #check if registry key exists to determine if form needs to be displayed\
    If( ((Get-ItemProperty "$RegHive\$RegPath" -Name Status -ErrorAction SilentlyContinue).Status -eq "Failed") ){
        Write-Verbose ("Last attempt failed; selector will be displayed")
        Start-TimeSelectorUI -UIObject $TZSelectUI -UpdateStatusKey "$RegHive\$RegPath"
    }
    ElseIf(-not((Get-ItemProperty "$RegHive\$RegPath" -Name Status -ErrorAction SilentlyContinue).Status -eq "Completed") ){
        Write-Verbose ("Last attempt did not complete; selector will be displayed")
        Start-TimeSelectorUI -UIObject $TZSelectUI -UpdateStatusKey "$RegHive\$RegPath"
    }
    Else{
        #do nothing
        #Stop-TimeSelectorUI -UIObject $TZSelectUI -UpdateStatusKey "$RegHive\$RegPath"-CustomStatus "Completed"
        Write-Verbose ("Selector has already ran once. Try -ForceTimeSelection param to run again.")
    }
}
Else{
    Write-Verbose ("All scenarios are false; selector will be displayed")
    Start-TimeSelectorUI -UIObject $TZSelectUI -UpdateStatusKey "$RegHive\$RegPath"
}
