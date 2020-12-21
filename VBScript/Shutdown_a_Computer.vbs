strComputer = "." ' Local Computer
Set objWMIService = GetObject("winmgmts:{impersonationLevel=impersonate,(Shutdown)}\\" & _
        		strComputer & "\root\cimv2")

Set colOs = objWMIService.ExecQuery("Select * from Win32_OperatingSystem")
 
For Each objOs in colOs
	objOs.Win32Shutdown(1)
Next