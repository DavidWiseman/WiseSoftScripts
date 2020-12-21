set objNetwork = createobject("wscript.network")

strADsPath = getUser(objNetwork.Username)

set objUser = GetObject(strADsPath)

wscript.echo "First Name: " & objUser.givenName & vbcrlf & _
	     "Surname: " & objUser.sn & vbcrlf & _
	     "Initials: " & objUser.initials & vbcrlf & _
	     "Display Name: " & objUser.displayName & vbcrlf & _
	     "Description: " & objUser.description & vbcrlf & _
	     "Office: " & objUser.physicalDeliveryOfficeName & vbcrlf & _
	     "Telephone Number: " & objUser.telephoneNumber & vbcrlf & _
	     "Email: " & objUser.mail & vbcrlf & _
	     "Web Page: " & objUser.wWWHomePage


Function getUser(Byval UserName)
	' Function to return the ADsPath from a username (sAMAccountName)
	' e.g. LDAP://cn=user1,cn=users,dc=wisesoft,dc=co,dc=uk

	DIM objRoot
	DIM getUserCn,getUserCmd,getUserRS

	on error resume next
	set objRoot = getobject("LDAP://RootDSE")

	set getUserCn = createobject("ADODB.Connection")
	set getUserCmd = createobject("ADODB.Command")
	set getUserRS = createobject("ADODB.Recordset")

	getUserCn.open "Provider=ADsDSOObject;"
	
	getUserCmd.activeconnection=getUserCn
	getUserCmd.commandtext="<LDAP://" & objRoot.get("defaultNamingContext") & ">;" & _
			"(&(objectCategory=user)(sAMAccountName=" & username & "));" & _
			"adsPath;subtree"


	
	set getUserRs = getUserCmd.execute

	
	if not rs.BOF and not rs.EOF then
     		getUserRs.MoveFirst
     		getUser = getUserRs(0)
	else
		getUser = ""
	end if

	getUserCn.close
end function