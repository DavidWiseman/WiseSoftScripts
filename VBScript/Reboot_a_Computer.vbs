strComputer = "." ' Local Computer

Set objWMIService = GetObject("winmgmts:{impersonationLevel=impersonate,(Shutdown)}!\\" & _
			strComputer & "\root\cimv2")

Set colOS = objWMIService.ExecQuery("Select * from Win32_OperatingSystem")

For Each objOS in colOS
	objOS.Reboot()
Next
	