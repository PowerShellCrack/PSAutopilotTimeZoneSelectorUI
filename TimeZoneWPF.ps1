
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
        Set to true if the deployment is for  autopilot 
        NOTE: Permission required for HKLM
    
    .PARAMETER TimeSelectorRunOnce
        Specify that this script will only launch the form one time.

    .PARAMETER ForceTimeSelection
        Disabled and with Bing API --> Current timezone and geo timezone will be compared; if different, form will be displayed
        Enabled --> the selection will always show

    .PARAMETER AutoTimeSelection
        Enabled with Bing API --> No prompt for user, time will update on it own
        Enabled without Bing API --> User will be prompted at least once
        Ignored if ForceTimeSelection is enabled

    .EXAMPLE
        PS> .\TimeZoneWPF.ps1 -IpStackAPIKey = "4bd1443445dfhrrt9dvefr45341" -BingMapsAPIKey = "Bh53uNUOwg71czosmd73hKfdHf465ddfhrtpiohvknlkewufjf4-d" -Verbose

        Uses IP GEO location for the pre-selection

    .EXAMPLE
        PS> .\TimeZoneWPF.ps1 -ForceTimeSelection

        This will always display the time selection screen; if IPStack and BingMapsAPI included the IP GEO location timezone will be preselected

    .EXAMPLE
        PS> .\TimeZoneWPF.ps1 -IpStackAPIKey = "4bd1443445dfhrrt9dvefr45341" -BingMapsAPIKey = "Bh53uNUOwg71czosmd73hKfdHf465ddfhrtpiohvknlkewufjf4-d" -AutoTimeSelection

        This will set the time automatically using the IP GEO location without prompting user. If API not provided, time will not change the time

    .EXAMPLE
        PS> .\TimeZoneWPF.ps1 -UserDriven $false

        Writes a registry key in HKLM hive to determine run status

    .EXAMPLE
        PS> .\TimeZoneWPF.ps1 -TimeSelectorRunOnce $true

        Mainly for Autopilot powershell scripts; this allows the screen to display one time after ESP is completed. 
#>

#===========================================================================
# CONTROL VARIABLES
#===========================================================================

[CmdletBinding(SupportsShouldProcess=$True)]
param(
    [string]$IpStackAPIKey = "",
    
    [string]$BingMapsAPIKey = "" ,

    [boolean]$UserDriven = $true,

    [boolean]$TimeSelectorRunOnce = $true,

    [switch]$ForceTimeSelection,
    
    [switch]$AutoTimeSelection
)

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

#replace some defualt attributes to support powershell
$inputXML = $inputXML -replace 'mc:Ignorable="d"','' -replace "x:N",'N'  -replace '^<Win.*', '<Window'

[void][System.Reflection.Assembly]::LoadWithPartialName('presentationframework')
[xml]$XAML = $inputXML
#Read XAML
$reader=(New-Object System.Xml.XmlNodeReader $xaml) 
  try{$Form=[Windows.Markup.XamlReader]::Load( $reader )}
catch{Write-Host "Unable to load Windows.Markup.XamlReader. Double-check syntax and ensure .net is installed."}

#===========================================================================
# Store Form Objects In PowerShell
#===========================================================================
#take the xaml properties and make them variables
$xaml.SelectNodes("//*[@Name]") | %{Set-Variable -Name "WPF$($_.Name)" -Value $Form.FindName($_.Name)}

Function Get-FormVariables{
    if ($global:ReadmeDisplay -ne $true){
        Write-host "To reference this display again, run Get-FormVariables" -ForegroundColor Yellow;
        $global:ReadmeDisplay=$true
    }
    write-host "Found the following interactable elements from our form" -ForegroundColor Cyan
    get-variable WPF*
}

#Get-FormVariables

#Set registry hive for user or local machine
If($UserDriven){$RegHive = 'HKCU:'}Else{$RegHive = 'HKLM:'}

# Build registry key for status and selection
#if unable to create key, deployment or permission may need to change 
Try{
    If(-not(Test-Path "$RegHive\SOFTWARE\TimezoneSelector") ){
        New-ItemProperty -Path "$RegHive\SOFTWARE" -nAME "TimezoneSelector" -ErrorAction Stop -Verbose | Out-Null       
    }
}
Catch{
    Throw ("Unable to configure registry key [{0}\{1}]. {3}" -f "$RegHive\SOFTWARE\TimezoneSelector", 'TimeZoneSelected ',$TargetTimeZone.id,$_.Exception.Message)
    Exit -1
}

#===========================================================================
# Actually make the objects work
#===========================================================================
#grab all timezones and add to list
function Get-GEOTimeZone {
    param(
        [CmdletBinding()]
        [string]$IpStackAPIKey,
        [string]$BingMapsAPIKey,
        [boolean]$AttemptOnline
    )
    Begin{
        ## Get the name of this function
        [string]${CmdletName} = $PSCmdlet.MyInvocation.MyCommand.Name

        if ($PSBoundParameters.ContainsKey('Verbose')) {
            $VerbosePreference = $PSCmdlet.SessionState.PSVariable.GetValue('VerbosePreference')
        }
    }
    Process{
        If($AttemptOnline){
            Write-Verbose "Attempting to check online for timezone"
            Write-Verbose "IPStack API: $IpStackAPIKey"
            Write-Verbose "Bing Maps API: $BingMapsAPIKey"

            $intuneManagementExtensionLogPath = "$env:ProgramData\Microsoft\IntuneManagementExtension\Logs\IntuneManagementExtension.log"

            #grab public IP and its geo location
            try {
                $geoIP = Invoke-RestMethod -Uri "http://api.ipstack.com/check?access_key=$($IpStackAPIKey)" -ErrorAction Stop -ErrorVariable $ErrorGeoIP
                Write-Verbose "Detected that $($geoIP.ip) is located in $($geoIP.country_name) at $($geoIP.latitude),$($geoIP.longitude)"
            }
            Catch {
                Write-Verbose "Error obtaining coordinates or public IP address" 
            }
            Finally{
                # Hide the api keys from logs to prevent manipulation API's
                (Get-Content -Path $intuneManagementExtensionLogPath).replace($IpStackAPIKey,'<sensitive data>') | Set-Content -Path $intuneManagementExtensionLogPath -ErrorAction SilentlyContinue | Out-Null
            }

            #determine geo location's timezone
            try {
                $timeZone = Invoke-RestMethod -Uri "https://dev.virtualearth.net/REST/v1/timezone/$($geoIP.latitude),$($geoIP.longitude)?key=$($BingMapsAPIKey)" -ErrorAction Stop -ErrorVariable $ErrortimeZone  
            }
            catch {
                Write-Verbose "Error obtaining Timezone from Bing Maps API"
            }
            Finally{
                # Hide the api keys from logs to prevent manipulation API's
                (Get-Content -Path $intuneManagementExtensionLogPath).replace($BingMapsAPIKey,'<sensitive data>')| Set-Content -Path $intuneManagementExtensionLogPath -ErrorAction SilentlyContinue | Out-Null
            }

            #if above worked, get selected time
            $correctTimeZone = $timeZone.resourceSets.resources.timeZone.windowsTimeZoneId
            Write-Verbose "Detected Correct time zone as '$($correctTimeZone)'"
            If($correctTimeZone){$SelectedTimeZone = [string](Get-TimeZone -id $correctTimeZone).DisplayName}

        }

        #confirm if time zone value exists, if not default to current time
        If(!$SelectedTimeZone){
            $SelectedTimeZone = [string](Get-TimeZone).DisplayName
            #$TargetTime = '(UTC-08:00) Pacific Time (US & Canada)'  
        }
    }
    End{
        #return selected timezone 
        return (Get-TimeZone -ListAvailable | Where {$_.Displayname -eq $SelectedTimeZone})
    }
}


#Get all timezones and load it to combo box
(Get-TimeZone -ListAvailable).DisplayName | ForEach-object {$WPFtargetTZ_listBox.Items.Add($_)} | Out-Null

#find a time zone to select

#splat if values exist
If( ([string]::IsNullOrEmpty($IpStackAPIKey)) -or ([string]::IsNullOrEmpty($BingMapsAPIKey)) ){
    $WPFtargetTZ_label.Text = $WPFtargetTZ_label.Text -replace "@anchor","What time zone are you in?"
    $params = @{
        AttemptOnline=$false
        Verbose=$VerbosePreference
    }
}
Else{
    $WPFtargetTZ_label.Text = $WPFtargetTZ_label.Text -replace "@anchor","Is this the time zone your in?"
    $params = @{
        AttemptOnline=$true
        ipStackAPIKey=$IpStackAPIKey
        bingMapsAPIKey=$BingMapsAPIKey
        Verbose=$VerbosePreference
    }
}

#grab Geo Timezone
$TargetGEOTimeZone = Get-GEOTimeZone @params

#select current time zone
$WPFtargetTZ_listBox.SelectedItem = $TargetGEOTimeZone.Displayname

#scrolls list to current selected item
#+3 below to center selected item on screen
$WPFtargetTZ_listBox.ScrollIntoView($WPFtargetTZ_listBox.Items[$WPFtargetTZ_listBox.SelectedIndex+3])

#if autoselection is enabled, attempt setting the time zone
If($AutoTimeSelection){
    Write-Verbose "Auto Selection enabled"
    Write-Verbose ("Attempting to auto set Time Zone to: {0}" -f $TargetGEOTimeZone.id)
    Set-TimeZone $TargetGEOTimeZone.id
    Start-Service W32Time | Restart-Service -ErrorAction SilentlyContinue
}

#compare the GEO Targeted timezone verses the current timezone
If($TargetGEOTimeZone.id -eq ((Get-TimeZone).Id) ){
    $TimeComparisonDiffers = $false
}
Else{
    $TimeComparisonDiffers = $true
}

#when button is clicked changer time
$WPFChangeTZButton.Add_Click({
    #Set time zone
    Set-TimeZone $TargetGEOTimeZone.id

    Write-Verbose ("Time Zone set: {0}" -f $TargetGEOTimeZone.id)
    #build registry key for time selector
    Set-ItemProperty -Path "$RegHive\SOFTWARE\TimezoneSelector" -Name TimeZoneSelected -Value "$($TargetGEOTimeZone.id)" -Force -ErrorAction Stop | Out-Null

    Get-Service W32Time | Restart-Service -ErrorAction SilentlyContinue
    Stop-TimeSelectorForm -StatusHive $RegHive})

#====================
# Shows the form
#====================
function Start-TimeSelectorForm{
    param(
        [CmdletBinding()]
        [string]$StatusHive
    )
    Set-ItemProperty -Path "$StatusHive\SOFTWARE\TimezoneSelector" -Name Status -Value "Running" -Force -ErrorAction SilentlyContinue | Out-Null
    
    Try{
        $Form.ShowDialog() | Out-Null
    }
    Catch{
        Set-ItemProperty -Path "$StatusHive\SOFTWARE\TimezoneSelector" -Name Status -Value 'Failed' -Force -ErrorAction Stop | Out-Null
    }
}

function Stop-TimeSelectorForm{
    param(
        [CmdletBinding()]
        [string]$StatusHive,
        [string]$CustomStatus
    )
    
    If($CustomStatus){$status = $CustomStatus}
    Else{$status = 'Completed'}

    Set-ItemProperty -Path "$StatusHive\SOFTWARE\TimezoneSelector" -Name Status -Value $status -Force -ErrorAction Stop | Out-Null
    $Form.Close() | Out-Null
}

#===========================================================================
# Main - Call the form depending on scneario
#===========================================================================

# found that if script is called by Intune, the script may be running multiple times if the ESP screen process takes a while
# Only allow the script to run once if it is already being displayed
If($ForceTimeSelection){
    #run form all the time
    Write-Verbose ("'Force Selection' parameter called: Form will be displayed")
    Start-TimeSelectorForm -StatusHive $RegHive
}
ElseIf((Get-ItemProperty "$RegHive\SOFTWARE\TimezoneSelector" -Name Status).Status -eq "Running"){
    Write-Verbose "Detected that TimeSelector form is running. Exiting"
    Exit
}
ElseIf($TimeComparisonDiffers -eq $true){
    #Only run if time compared differs
    Write-Verbose ("Current time is different than Geo time scenario: Form will be displayed")
    Start-TimeSelectorForm -StatusHive $RegHive
}
ElseIf($TimeSelectorRunOnce){
    #check if regsitry key exists to determine if form needs to be displayed\
    If(-not((Get-ItemProperty "$RegHive\SOFTWARE\TimezoneSelector" -Name Status).Status -eq "Completed") ){
        Write-Verbose ("No key exist for run once scenario: Form will be displayed")
        Start-TimeSelectorForm -StatusHive $RegHive
    }
    Else{
        #do nothing
        Stop-TimeSelectorForm -StatusHive $RegHive -CustomStatus "Completed"
        Return
    }
}
Else{
    Write-Verbose ("All scenarios are false: Form will be displayed")
    Start-TimeSelectorForm -StatusHive $RegHive
}
