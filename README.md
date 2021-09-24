# Select Time Zone

## Overview
I wrote this UI in PowerShell for devices to users to change the time zone on the device. This was originally intended for a user to select their timezone on first logon. I wanted the this to popup to look and feel like something from Windows Out-of-Box-Experience (OOBE).

![Alt_text](https://1.bp.blogspot.com/-4A1HjymzCik/XfWWtuKcFTI/AAAAAAAAWJM/Ae36IYLmsIAOXQl4PP9wHvdDYZbkovPAgCLcBGAsYHQ/s1600/Set-timeone.PNG)

![Alt_text](https://2.bp.blogspot.com/-Ni2wM-CS7ik/XfWWcTREKDI/AAAAAAAAWJE/J8y1FdYBCeMG-2TaPHDJIivHmnixMYpLwCLcBGAsYHQ/s1600/Set-timeone_select.PNG)


## Read to deploy?

### For Autopilot
- This is only tested on a single user device. have not tested it on multiuser or Kiosk device.
- This runs in User context but for the device


1. Login into endpoint.microsoft.com
1. Navigate to Devices-->Scripts
1. Click _Add_ --> Windows 10 and Later
1. give it a name (eg. Win10 TimeZone Selector)
1. Import the PowerShell script
1. Select **Yes** for Run this script using the logged on credentials
1. Select **No** for Enforce script signature check
1. Select **No** for Run script in 64 bit PowerShell Host
1. Build an Azure Dynamic Device Group using query:

```kusto
    (device.devicePhysicalIDs -any _ -contains "[ZTDID]")
```

1. Assign script to Azure Dynamic Device Group


# DISCLAIMER

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS 
OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, 
FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE 
AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER 
LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING 
FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER 
DEALINGS IN THE SOFTWARE.
