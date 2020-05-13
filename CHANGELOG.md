# Change log for TimeSelectorWPF.ps1

## Unreleased

## 1.5.2 - Apr 29, 2020
 - changed main script variables to parameters
 - moved changelog to markdown
 - Updated help 

## 1.5.1 - Apr 29, 2020
 - Fixed TimeComparisonDiffers to Check if true (not false)

## 1.5.0 - Apr 10, 2020
 - Removed OS image check and used registry key. 
 - Set values for verbose logging
 - Set user driven mode to merge autopilot scenario 
 
## 1.4.2 - Mar 26, 2020
 - Remove API key from Intune management log for sensitivity 
 
## 1.4.1 - Mar 06, 2020
 - Check if Select time found when no API specified
 
## 1.4.0 - Feb 13, 2020
 - Added the online geo time check to select appropiate time based on public IP;Requires API keys for Bingmaps and ipstack.com
 
## 1.3.0 - Jan 21, 2020
 - Removed Time Zone select notification
 - increase height timzone list
 
## 1.2.8 - Jan 16, 2020
 - Scrolls to current time zone
 
## 1.2.6 - Dec 19, 2019
 - Added image date checker for AutoPilot scenarios; won't launch form if not imaged within 2 hours
 
## 1.2.5 - Dec 19, 2019
 - Centered grid to support different resolutions
 - changed font to light
 
## 1.2.1 - Dec 16, 2019
 - Highlighted current timezne in yellow
 - centered text in grid columns
 
## 1.2.0 - Dec 14, 2019
 - Styled theme to look like OOBE
 - changed Combobox to ListBox
 
## 1.1.0 - Dec 12, 2019
 - Centered all lines
 - changed background
 
## 1.0.0 - Dec 09, 2019
 - initial