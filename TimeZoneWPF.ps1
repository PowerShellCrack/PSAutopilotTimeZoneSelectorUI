
<#

    .SYNOPSIS
        Prompts user to set time zone
    
    .DESCRIPTION
		Prompts user to set time zone using windows presentation framework
        Can be used in:
            - SCCM Tasksequences (User interface allowed)
            - SCCM Software Delivery (User interface allowed)
            - Intune Autopilot 
   
    .INFO
        Author:         Richard "Dick" Tracy II
        Last Update:    03/26/2020
        Version:        1.5.0
        Thanks:         Eric Moe,Matthew White

    .NOTES
        Launches in full screen

    .CHANGE LOGS
        1.5.0 - Apr 10, 2020 - Remvoed OS image check and used registry key. Set values for verbose logging, and user deriven mode to merge autopilot version 
        1.4.2 - Mar 26, 2020 - Remove API key from Intune management log for sensitivity 
        1.4.1 - Mar 06, 2020 - Check if Select time found when no API specified
        1.4.0 - Feb 13, 2020 - Added the onloine geo time check to select appropiate time based on public IP;
                               Requires API keys for Bingmaps and ipstack.com
        1.3.0 - Jan 21, 2020 - Removed Time Zone select notification, increase height timzone list
        1.2.8 - Jan 16, 2020 - Scrolls to current time zone
        1.2.6 - Dec 19, 2019 - Added image date checker for AutoPilot scenarios; won't launch form if not imaged within 2 hours
        1.2.5 - Dec 19, 2019 - Centered grid to support different resolutions; changed font to light
        1.2.1 - Dec 16, 2019 - Highlighted current timezne in yellow; centered text in grid columns
        1.2.0 - Dec 14, 2019 - Styled theme to look like OOBE; changed Combobox to ListBox
        1.1.0 - Dec 12, 2019 - Centered all lines; changed background
        1.0.0 - Dec 09, 2019 - initial

    .LINKS
        https://matthewjwhite.co.uk/2019/04/18/intune-automatically-set-timezone-on-new-device-build/
        https://ipstack.com
        https://azuremarketplace.microsoft.com/en-us/marketplace/apps/bingmaps.mapapis
    -------------------------------------------------------------------------------
    LEGAL DISCLAIMER
    This Sample Code is provided for the purpose of illustration only and is not
    intended to be used in a production environment.  THIS SAMPLE CODE AND ANY
    RELATED INFORMATION ARE PROVIDED "AS IS" WITHOUT WARRANTY OF ANY KIND, EITHER
    EXPRESSED OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE IMPLIED WARRANTIES OF
    MERCHANTABILITY AND/OR FITNESS FOR A PARTICULAR PURPOSE.  We grant You a
    nonexclusive, royalty-free right to use and modify the Sample Code and to
    reproduce and distribute the object code form of the Sample Code, provided
    that You agree: (i) to not use Our name, logo, or trademarks to market Your
    software product in which the Sample Code is embedded; (ii) to include a valid
    copyright notice on Your software product in which the Sample Code is embedded;
    and (iii) to indemnify, hold harmless, and defend Us and Our suppliers from and
    against any claims or lawsuits, including attorneys’ fees, that arise or result
    from the use or distribution of the Sample Code.
 
    This posting is provided "AS IS" with no warranties, and confers no rights. Use
    of included script samples are subject to the terms specified
    at https://www.microsoft.com/en-us/legal/copyright/.
    -----------------------------------------------------------------------------

#>

#===========================================================================
# CONTROL VARIABLES
#===========================================================================
#used to get geoCoordinates of the public IP. get the API key from https://ipstack.com
$ipStackAPIKey = "4bd144c23e13947562b73ca8644aa431" 

#Used to get the Windows TimeZone value of the location coordinates. get the API key from https://azuremarketplace.microsoft.com/en-us/marketplace/apps/bingmaps.mapapis
$bingMapsAPIKey = "An19uNUOwg71czomO2cEB9njocBF9Ip7SV82Kmp6Fkg_Gk6VLTMc6tXGuwbAs8-f" 

# deploy to user sets either HKCU key or HKLM key
# Set to true if the deployment is for  autopilot 
# NOTE: Permission required for HKLM
$UserDriven = $true

# Specify that this script will only launch the form one time.
$RunTimeSelectorOnce = $true

#$VerbosePreference = 'Continue'
$VerbosePreference = 'SilentlyContinue'
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

#===========================================================================
# Actually make the objects work
#===========================================================================
#get the current timezone and display it in UI
#$DefaultTime = (Get-TimeZone).DisplayName
#$WPFCurrentTZ.Text = $WPFCurrentTZ.Text -replace "@anchor",$DefaultTime

#grab all timezones and add to list
function Get-SelectedTime {
    param(
        [CmdletBinding()]
        [string]$ipStackAPIKey,
        [string]$bingMapsAPIKey,
        [switch]$AttemptOnline
    )
    Begin{
        ## Get the name of this function
        [string]${CmdletName} = $PSCmdlet.MyInvocation.MyCommand.Name

        if ($PSBoundParameters.ContainsKey('Verbose')) {
            $VerbosePreference = $PSCmdlet.SessionState.PSVariable.GetValue('VerbosePreference')
        }
    }
    Process{
        If($PSBoundParameters.ContainsKey('AttemptOnline')){
            Write-Verbose "Attempting to check online for timezone"
            Write-Verbose "IPStack API: $ipStackAPIKey"
            Write-Verbose "Bing Maps API: $bingMapsAPIKey"

            # Hide the api keys from logs to prevent manipulation API's
            # This area was designed for Intune Managment Extension
            $intuneManagementExtensionLogPath = "$env:ProgramData\Microsoft\IntuneManagementExtension\Logs\IntuneManagementExtension.log"
            (Get-Content -Path $intuneManagementExtensionLogPath).replace($ipStackAPIKey,'<sensitive data>')| Set-Content -Path $intuneManagementExtensionLogPath -ErrorAction SilentlyContinue | Out-Null
            (Get-Content -Path $intuneManagementExtensionLogPath).replace($bingMapsAPIKey,'<sensitive data>')| Set-Content -Path $intuneManagementExtensionLogPath -ErrorAction SilentlyContinue | Out-Null

            #grab public IP and its geo location
            try {
                $geoIP = Invoke-RestMethod -Uri "http://api.ipstack.com/check?access_key=$($ipStackAPIKey)" -ErrorAction Stop -ErrorVariable $ErrorGeoIP
                Write-Verbose "Detected that $($geoIP.ip) is located in $($geoIP.country_name) at $($geoIP.latitude),$($geoIP.longitude)"
            }
            Catch {
                Write-Verbose "Error obtaining coordinates or public IP address" 
            }

            #determine geo location's timezone
            try {
                $timeZone = Invoke-RestMethod -Uri "https://dev.virtualearth.net/REST/v1/timezone/$($geoIP.latitude),$($geoIP.longitude)?key=$($bingMapsAPIKey)" -ErrorAction Stop -ErrorVariable $ErrortimeZone
            }
            catch {
                Write-Verbose "Error obtaining Timezone from Bing Maps API"
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
If([string]::IsNullOrEmpty($ipStackAPIKey) -and [string]::IsNullOrEmpty($bingMapsAPIKey)){
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
        ipStackAPIKey=$ipStackAPIKey
        bingMapsAPIKey=$bingMapsAPIKey
        Verbose=$VerbosePreference
    }
}
$TargetTimeZone = Get-SelectedTime @params

#select current time zone
$WPFtargetTZ_listBox.SelectedItem = $TargetTimeZone.Displayname

#scrolls list to current selected item
#+3 below to center selected item on screen
$WPFtargetTZ_listBox.ScrollIntoView($WPFtargetTZ_listBox.Items[$WPFtargetTZ_listBox.SelectedIndex+3])


#when button is clicked changer time
$WPFChangeTZButton.Add_Click({
    #Set time zone
    Set-TimeZone $TargetTimeZone.id

    Write-Verbose ("Time Zone set: {0}" -f $TargetTimeZone.id)
    #build registry key for time selector
    Try{
        If(-not(Test-Path "$RegHive\SOFTWARE\TimezoneSelector") ){
            New-Item "$RegHive\SOFTWARE\TimezoneSelector" -Force -ErrorAction Stop | New-ItemProperty -Name TimeZoneSelected -PropertyType String -Value "$($TargetTimeZone.id)" -Force -ErrorAction Stop -Verbose:$VerbosePreference | Out-Null       
        } 
        Else{
            Set-ItemProperty -Path "$RegHive\SOFTWARE\TimezoneSelector" -Name TimeZoneSelected -Value "$($TargetTimeZone.id)" -Force -ErrorAction Stop -Verbose:$VerbosePreference | Out-Null
        }
    }
    Catch{
        Write-Error ("Unable to configure registry key [{0}\{1}]. {3}" -f "$RegHive\SOFTWARE\TimezoneSelector", 'TimeZoneSelected ',$TargetTimeZone.id,$_.Exception.Message)
    }

    Start-Service W32Time -ErrorAction SilentlyContinue | Restart-Service -ErrorAction SilentlyContinue -Verbose:$VerbosePreference
    $Form.Close()})

#====================
# Shows the form
#====================
function Show-Form{
    $Form.ShowDialog() | out-null
}


#===========================================================================
# Main - Call the form
#===========================================================================
If($RunTimeSelectorOnce){
    #check if regsitry key exists to determine if form needs to be displayed
    If(-not(Test-Path "$RegHive\SOFTWARE\TimezoneSelector")){
        Show-Form
    }Else{
        #do nothing
        Return
    }
}
Else{
   #run form all the time
   Show-Form
}
