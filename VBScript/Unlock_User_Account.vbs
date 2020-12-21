username=inputbox("Enter username:")
if username = "" then wscript.quit

ldapPath = FindUser(username)

if ldapPath = "Not Found" then
	wscript.echo "User not found!"
else
	set objUser = getobject(ldapPath)
	if isAccountLocked(objUser) then
		objuser.put "lockoutTime", 0
		objUser.setinfo
		wscript.echo "Account Unlocked"
	else
		wscript.echo "This account is not locked out"
	end if
end if


Function FindUser(Byval UserName) 
	on error resume next

	set objRoot = getobject("LDAP://RootDSE")
	domainName = objRoot.get("defaultNamingContext")
	set cn = createobject("ADODB.Connection")
	set cmd = createobject("ADODB.Command")
	set rs = createobject("ADODB.Recordset")

	cn.open "Provider=ADsDSOObject;"
	
	cmd.activeconnection=cn
	cmd.commandtext="SELECT ADsPath FROM 'LDAP://" & domainName & _
			"' WHERE sAMAccountName = '" & UserName & "'"
	
	set rs = cmd.execute

	if err<>0 then
		wscript.echo "Error connecting to Active Directory Database:" & err.description
		wscript.quit
	else
		if not rs.BOF and not rs.EOF then
     			rs.MoveFirst
     			FindUser = rs(0)
		else
			FindUser = "Not Found"
		end if
	end if
	cn.close
end function

Function IsAccountLocked(byval objUser)
    	on error resume next
	set objLockout = objUser.get("lockouttime")

	if err.number = -2147463155 then
		isAccountLocked = False
		exit Function
	end if
	on error goto 0
	
	if objLockout.lowpart = 0 And objLockout.highpart = 0 Then
		isAccountLocked = False
	Else
		isAccountLocked = True
	End If

End Function