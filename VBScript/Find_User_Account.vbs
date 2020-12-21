domainname=inputbox("Enter DNS Domain Name" & vbcrlf & "(Leave blank for current domain):")
username=inputbox("Enter username:")


if domainname = "" then
	set objRoot = getobject("LDAP://RootDSE")
	domainname = objRoot.get("defaultNamingContext")
end if

if username <> "" then
	wscript.echo finduser(username,domainname)
end if


Function FindUser(Byval UserName, Byval Domain) 
	on error resume next

	set cn = createobject("ADODB.Connection")
	set cmd = createobject("ADODB.Command")
	set rs = createobject("ADODB.Recordset")

	cn.open "Provider=ADsDSOObject;"
	
	cmd.activeconnection=cn
	cmd.commandtext="SELECT ADsPath FROM 'LDAP://" & Domain & _
			 "' WHERE sAMAccountName = '" & UserName & "'"
	
	set rs = cmd.execute

	if err<>0 then
		FindUser="Error connecting to Active Directory Database:" & err.description
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