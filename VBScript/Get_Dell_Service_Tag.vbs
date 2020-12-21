strComputer = "." 

Set objWMIService = GetObject("winmgmts:" & _
		"{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2") 

For Each objSMBIOS in objWMIService.ExecQuery("Select * from Win32_SystemEnclosure") 
	Wscript.Echo "Serial Number: " & objSMBIOS.SerialNumber 
Next
