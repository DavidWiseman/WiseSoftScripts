option explicit
DIM objNetwork, strComputer
DIM strPassword, strAdminUserName, strNewAdminUserName

set objNetwork = CreateObject("Wscript.Network")
strComputer = UCASE(objNetwork.ComputerName)

' The old name of the administrator user account (normally administrator)
strAdminUserName = "Administrator"
' The new name of the administrator user account
strNewAdminUserName = "NormalUser"
' Password includes computername to have a unique password on all computers.
strPassword = "PrefixSTDP@$$w0rd" & strComputer

' Rename admin user account
renameUser strComputer,strAdminUserName,strNewAdminUserName
' Set password of admin user account
setPWD strComputer,strNewAdminUserName,strPassword

' Reset password for a local user account on a given computer
sub setPWD(strComputer,strUser,strPassword)

	DIM objUser
	' Ignore error if user account isn't found or error changing password
	on error resume next 
	set objUser = getobject("WinNT://" & strComputer & "/" & strUser & ",user")
	if err.number = 0 then
		objUser.SetPassword strPassword
		objUser.SetInfo
	end if
	on error goto 0

end sub

' Rename a local user account on a given computer
sub renameUser(strComputer,strFromName, strToName)
	
	DIM objComputer,objUser
	' Ignore error if user account isn't found or error moving user
	on error resume next
	set objComputer = GetObject("WinNT://" & strComputer)
	set objUser = getobject("WinNT://" & strComputer & "/" & strFromName & ",user")
	if err.number = 0 then
		objComputer.MoveHere objUser.ADsPath,strToName
	end if
	on error goto 0
	
end sub