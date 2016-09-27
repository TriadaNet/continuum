$strComputer = "."

gwmi WmiMonitorID -Namespace root\wmi -computername $strComputer |
   Select PSComputerName,
     @{n="Model";e={[System.Text.Encoding]::ASCII.GetString($_.UserFriendlyName -ne 00)}},
     @{n="Serial Number";e={[System.Text.Encoding]::ASCII.GetString($_.SerialNumberID -ne 00)}} |
     Format-List | Out-File “c:\!installs\Monitors.txt”

gwmi win32_videocontroller -computername $strComputer | select caption, CurrentHorizontalResolution, CurrentVerticalResolution | fl | Out-File -Append “c:\!installs\Monitors.txt”
gwmi win32_desktopmonitor -computername $strComputer | fl MonitorManufacturer,Name | Out-File -Append “c:\!installs\Monitors.txt”


     
