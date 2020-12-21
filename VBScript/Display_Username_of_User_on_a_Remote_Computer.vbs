strComputer = inputbox("Please enter the name of the computer:")

' Check that user entered a value
if strComputer = "" then
	wscript.quit
end if

ON ERROR RESUME NEXT ' Handle errors connecting to the computer (Not switched on, permissions error etc)
set objWMI = GetObject("winmgmts:{impersonationLevel=impersonate}!//" & strComputer & "")

if err <> 0 then ' Check for error
	wscript.echo "Error connecting to specified computer: " & err.description
	wscript.quit
end if
ON ERROR GOTO 0 ' Turn off resume next error handling

set colOS = objWMI.ExecQuery("Select * from Win32_ComputerSystem")

For Each objItem In colOS
	if strUsers <> "" then
		strUsers = strUsers & ", " & objItem.UserName
	else
		strUsers = objItem.UserName
	End If
Next

wscript.echo "The following user(s) are logged on to " & strComputer & ":" & strUsers
