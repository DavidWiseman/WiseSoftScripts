strComputer = "."
strProcessToKill = "iexplore.exe" 

Set objWMIService = GetObject("winmgmts:" _
	& "{impersonationLevel=impersonate}!\\" _ 
	& strComputer & "\root\cimv2") 

Set colProcess = objWMIService.ExecQuery _
	("Select * from Win32_Process Where Name = '" & strProcessToKill & "'")

count = 0
For Each objProcess in colProcess
	objProcess.Terminate()
	count = count + 1
Next 


wscript.echo "Killed " & count & " instances of " & _
	strProcessToKill & "on " & strComputer


