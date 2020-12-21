strComputer = "." ' Local computer
strMemory = ""
i = 1
      
set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")
set colItems = objWMIService.ExecQuery("Select * from Win32_PhysicalMemory")

For Each objItem In colItems

	if strMemory <> "" then strMemory = strMemory & vbcrlf
	strMemory = strMemory &  "Bank" & i & " : " & (objItem.Capacity / 1048576) & " Mb"
	i = i + 1
Next
installedModules = i - 1

Set colItems = objWMIService.ExecQuery("Select * from Win32_PhysicalMemoryArray")

For Each objItem in colItems
	totalSlots = objItem.MemoryDevices
Next
	

wscript.echo "Total Slots: " & totalSlots & vbcrlf & _
	     "Free Slots: " & (totalSlots - installedModules) & vbcrlf & _
	     vbcrlf & "Installed Modules:" & vbcrlf & strMemory
