$strComputer = "."
$colSlots = Get-WmiObject -Class "win32_PhysicalMemoryArray" -namespace "root\CIMV2"  -computerName $strComputer
$colRAM = Get-WmiObject -Class "win32_PhysicalMemory" -namespace "root\CIMV2"  -computerName $strComputer
"Computer Name :" + $env:computername  | Out-File “c:\!installs\Memory.txt”

Foreach ($objSlot In $colSlots){
    "Total Number of DIMM Slots: " + $objSlot.MemoryDevices | Out-File -Append “c:\!installs\Memory.txt”
}
Foreach ($objRAM In $colRAM) {
    "Memory Installed: " + $objRAM.DeviceLocator | Out-File -Append “c:\!installs\Memory.txt”
    "Memory Size: " + ($objRAM.Capacity / 1GB) + " GB“ | Out-File -Append “c:\!installs\Memory.txt”
}
