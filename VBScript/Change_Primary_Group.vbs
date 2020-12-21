'<<<< Prompt for a username >>>>
userName = inputbox("Please enter a username:")
if username = "" then wscript.quit
'<<<< Prompt for a group name >>>>
groupName = inputBox("Please enter the name of a group:")
if GroupName = "" then wscript.quit

'<<<< Get the default naming context (saves hard-coding the domain name) >>>>
set objRoot = getobject("LDAP://RootDSE")
defaultNC = objRoot.get("defaultnamingcontext")

'<<<< Find the ADsPath of the user object >>>>
userDN = FindObject(userName,defaultNC,"User")

'<<<< Find the ADsPath of the group object >>>>
groupDN=findObject(groupName,defaultNC,"Group")

'<<<< Find the primaryGroupToken of the group >>>>
intPrimaryGroupToken = getPrimaryGroupToken(groupName,defaultNC)

'<<<< Quit if any query failed >>>>
if userDN = "Not Found" or GroupDN = "Not Found" or intPrimaryGroupToken = "Not Found" then
	wscript.echo "User or Group not found!"
	wscript.quit
end if

'<<<< Bind to the group object >>>>
set objGroup = getobject(groupDN)

'<<<< Bind to the user object >>>>
set objUser = getobject(userDN)

'<<<< Check if primary group is already set >>>>
if objuser.primarygroupid = intPrimaryGroupToken then
	wscript.echo "Primary Group already set to " & groupName
	wscript.quit
end if

'<<<< Add user to group >>>>
addToGroup objUser.adspath, objGroup.adspath

'<<<< Change the primary group >>>>
objUser.primaryGroupID = intPrimaryGroupToken
objUser.setinfo

wscript.echo "Finished"

'<<<< Returns an ADsPath of a given object >>>>
Function FindObject(Byval sAMAccountName, Byval searchRoot,Byval objectCategory) 
	on error resume next

	set cn = createobject("ADODB.Connection")
	set cmd = createobject("ADODB.Command")
	set rs = createobject("ADODB.Recordset")

	cn.open "Provider=ADsDSOObject;"
	
	cmd.activeconnection=cn
	cmd.commandtext="<LDAP://" & searchRoot & ">;" & _
		"(&(objectCategory=" & objectCategory & _
		 ")(sAMAccountName=" & sAMAccountName & "));adspath;subtree"

	set rs = cmd.execute

	if err<>0 then
		wscript.echo "Error connecting to Active Directory Database:" & err.description
		wscript.quit
	else
		if not rs.BOF and not rs.EOF then
     			rs.MoveFirst
     			FindObject = rs(0)
		else
			FindObject = "Not Found"
		end if
	end if
	cn.close
end function

'<<<< Returns the primaryGroupToken of a group >>>>
Function getPrimaryGroupToken(Byval sAMAccountName, Byval searchRoot) 
	on error resume next

	set cn = createobject("ADODB.Connection")
	set cmd = createobject("ADODB.Command")
	set rs = createobject("ADODB.Recordset")

	cn.open "Provider=ADsDSOObject;"
	
	cmd.activeconnection=cn
	cmd.commandtext="<LDAP://" & searchRoot & ">;" & _
		"(&(objectCategory=Group)(sAMAccountName=" & _
		 sAMAccountName & "));primaryGroupToken;subtree"

	set rs = cmd.execute

	if err<>0 then
		wscript.echo "Error connecting to Active Directory Database:" & err.description
		wscript.quit
	else
		if not rs.BOF and not rs.EOF then
     			rs.MoveFirst
     			getPrimaryGroupToken = rs("primaryGroupToken")
		else
			getPrimaryGroupToken = "Not Found"
		end if
	end if
	cn.close
end function

'<<<< Adds user to group >>>>
sub addToGroup(userDN, groupDN)
	
	set objGroup = getobject(groupDN)
	
	for each member in objGroup.members
		if lcase(member.adspath) = lcase(userdn) then
			exit sub
		wscript.echo "TEST"
		end if
	next
	objGroup.Add(userDN)

end sub