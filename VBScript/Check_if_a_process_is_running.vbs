option explicit
DIM strComputer,strProcess

strComputer = "." ' local computer
strProcess = "calc.exe"

' Check if Calculator is running on specified computer (. = local computer)
if isProcessRunning(strComputer,strProcess) then
	wscript.echo strProcess & " is running on computer '" & strComputer & "'"
else
	wscript.echo strProcess & " is NOT running on computer '" & strComputer & "'"
end if

' Function to check if a process is running
function isProcessRunning(byval strComputer,byval strProcessName)

	Dim objWMIService, strWMIQuery

	strWMIQuery = "Select * from Win32_Process where name like '" & strProcessName & "'"
	
	Set objWMIService = GetObject("winmgmts:" _
		& "{impersonationLevel=impersonate}!\\" _ 
			& strComputer & "\root\cimv2") 


	if objWMIService.ExecQuery(strWMIQuery).Count > 0 then
		isProcessRunning = true
	else
		isProcessRunning = false
	end if

end function