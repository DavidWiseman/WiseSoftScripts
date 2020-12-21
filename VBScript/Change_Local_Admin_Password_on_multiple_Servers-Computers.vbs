' Comma-seperated list of member servers (or client computers)
strServers = "server01,server02,server03,server04"
' Prompt user for password
strPassword = inputbox("Please enter a new password:")
' Check that user entered a password
if strPassword = "" then
	wscript.quit
end if
on error resume next
' Enumerate each server in the comma-sepereated list
for each strServer in split(strServers,",")
	' Connect to Administrator acccount on server using WinNT provider
	set objUser = getobject("WinNT://" & strServer & "/administrator,user")
	' Check if we connected to the user object successfully
	if err.Number <> 0 then
		' Display an error message & clear the error
		msgbox "Unable to connect to Administrator " & _
			     "user object on server " & strServer & ":" & _
			      vbcrlf & "Error #" & err.Number & vbcrlf & _
			      vbcrlf & err.Description,vbOkOnly+vbCritical,"Error"
		err.Clear
	else
		' Change the password
		objUser.SetPassword strPassword
		objUser.SetInfo ' Save Changes
		if err.Number <> 0 then
			' Display an error message & clear the error
			msgbox "Unable to change the Administrator password " & _
				     "on server " & strServer & ":" & vbcrlf & _
				     "Error #" & err.Number & vbcrlf & _
				     err.Description,vbOkOnly+vbCritical,"Error"
			err.Clear
		end if
	end if
next