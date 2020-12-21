option explicit

Dim strComputer,strServiceName
strComputer = "." ' Local Computer
strServiceName = "wuauserv" ' Windows Update Service

if isServiceRunning(strComputer,strServiceName) then
	wscript.echo "The '" & strServiceName & "' service is running on '" & strcomputer & "'"
else
	wscript.echo "The '" & strServiceName & "' service is NOT running on '" & strcomputer & "'"
end if

' Function to check if a service is running on a given computer
function isServiceRunning(strComputer,strServiceName)
	Dim objWMIService, strWMIQuery

	strWMIQuery = "Select * from Win32_Service Where Name = '" & strServiceName & "' and state='Running'"

	Set objWMIService = GetObject("winmgmts:" _
		& "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")

	if objWMIService.ExecQuery(strWMIQuery).Count > 0 then
		isServiceRunning = true
	else
		isServiceRunning = false
	end if

end function