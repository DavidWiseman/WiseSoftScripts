const FileName = "groupmembers.csv"

groupName = inputbox("Please enter the name of the group:")

if groupName = "" then
	wscript.quit
end if

groupPath = getgrouppath(groupName)

if groupPath = "" then
	wscript.echo "Unable to find the specified group in the domain"
	wscript.quit
end if

set objGroup = getobject(grouppath)
set objFSO = createobject("scripting.filesystemobject")
set objFile = objFSO.createtextfile(FileName)
q = """"


objFile.WriteLine(q & "sAMAccountName" & q & "," & q & "Surname" & q & "," & q & "FirstName" & q)
for each objMember in objGroup.Members
	objFile.WriteLine(q & objmember.samaccountname & q & "," & q & objmember.sn & _
			q & "," & q & objmember.givenName & q)
next

'***** Users who's primary group is set to the given group need to be enumerated seperatly.*****
getprimarygroupmembers groupname

objFile.Close

wscript.echo "Completed"

function getGroupPath(byval GroupName)
	set cmd=createobject("ADODB.Command")
	set cn=createobject("ADODB.Connection")
	set rs=createobject("ADODB.Recordset")
	
	cn.open "Provider=ADsDSOObject;"
	
	cmd.commandtext = "SELECT adspath from 'LDAP://" & getnc & _
			  "' WHERE objectCategory = 'Group' and sAMAccountName = '" & groupname & "'"
	cmd.activeconnection = cn
	
	set rs = cmd.execute
	
	if rs.bof <> true and rs.eof<>true then
		getgrouppath=rs(0)
	else
		getgrouppath = ""
	end if
	cn.close

end function

function getNC
	set objRoot=getobject("LDAP://RootDSE")
	getNC=objRoot.get("defaultNamingContext")
end function

function getPrimaryGroupMembers(byval GroupName)
	set cn = createobject("ADODB.Connection")
	set cmd = createobject("ADODB.Command")
	set rs = createobject("ADODB.Recordset")
	
	cn.open "Provider=ADsDSOObject;"
	cmd.activeconnection=cn

	'***** Change the Page Size to overcome the 1000 record limitation *****
	cmd.properties("page size")=1000
	cmd.commandtext = "SELECT PrimaryGroupToken FROM 'LDAP://" & getnc & _
			  "' WHERE sAMAccountName = '" & GroupName & "'"
	set rs = cmd.execute

	if rs.eof<>true and rs.bof<>true then
		PrimaryGroupID = rs(0)
	else
		Err.Raise 5000, "getPrimaryGroupMembers", "Unable to find PrimaryGroupToken property"
	end if

	cmd.commandtext = "SELECT samaccountname, sn, givenName FROM 'LDAP://" & getNC & _
			  "' WHERE PrimaryGroupID = '" & PrimaryGroupID & "'"

	set rs = cmd.execute

	while rs.eof<>true and rs.bof<>true
		objFile.WriteLine(q & rs("samaccountname") & q & "," & q & rs("sn") & q & _
				  "," & q & rs("givenName") & q)
		rs.movenext
	wend
	cn.close
	
end function