OPTION EXPLICIT
Const ADS_PROPERTY_CLEAR = 1
Const ADS_PROPERTY_DELETE = 4
Const E_ADS_PROPERTY_NOT_FOUND  = &h8000500D

DIM strSearchFilter, strSearchRoot, objRootDSE
DIM cn,cmd,rs, strSearchScope
DIM objNewOU,objUser, strNewOU, strNewGroups, strGroup
DIM i
DIM arrGroups()

' ********************************************************
' * Setup
' ********************************************************

' Specify the distinguished name of the new OU
strNewOU = "ou=Marketing,ou=All Users,dc=wisesoft,dc=org,dc=uk"

' Specify a list of new groups (Use a semi-colon to separate)
strNewGroups = "Marketing;Sales"

' Modify the filter to query for your department.  
' This filter will find all users in the marketing department
strSearchFilter = "(&(objectCategory=Person)(objectClass=User)(department=marketing))"

' Specify a search root. The domain root is used by default. 
' e.g. dc=wisesoft,dc=co,dc=uk
' You could also specify a particular OU to start the search from.
' e.g. strSearchRoot = "ou=students,ou=All Users,dc=wisesoft,dc=co,dc=uk"
strSearchRoot = getDomainRoot

' A value of "subtree" will search all child containers (OUs).
' Change to "onelevel" if you don't want child containers to be 
' included in the search
strSearchScope = "subtree"

' ********************************************************
set objNewOU = GetObject("LDAP://" & strNewOU)

Set cn = CreateObject("ADODB.Connection")
Set cmd =   CreateObject("ADODB.Command")
cn.open "Provider=ADsDSOObject;"

Set cmd.ActiveConnection = cn

cmd.CommandText = "<LDAP://" & strSearchRoot & ">;" & strSearchFilter & ";ADsPath;" & strSearchScope
cmd.Properties("Page Size") = 1000

Set rs = cmd.Execute


i=0

' Get the distinguished name for each of the new groups and store in the array
for each strGroup in SPLIT(strNewGroups,";")
	REDIM PRESERVE arrGroups(i)
	arrGroups(i) = GetGroupDN(strGroup)
	i = i + 1
next

' loop through the search results
while rs.eof<> true and rs.bof<>true
	set objUser = GetObject(rs(0))

	' Remove all existing groups except primary group
	clearGroupMembership objUser
	
	' Add new groups
	for each strGroup in arrGroups
		addToGroup objUser.Get("distinguishedName"), strGroup
	next

	' Move user to new ou (passing the ADsPath attribute returned from the query)
	objNewOU.MoveHere rs(0),vbNullString

	rs.movenext
wend

rs.close
cn.close

wscript.echo "Completed"

private function getDomainRoot
	' Bind to RootDSE - this object is used to 
	' get the default configuration naming context
	' e.g. dc=wisesoft,dc=co,dc=uk

	set objRootDSE = getobject("LDAP://RootDSE")
	getDomainRoot = objRootDSE.Get("DefaultNamingContext")
end function


private sub clearGroupMembership(byref objUser)
	' Clear all existing group membership (primary group is ignored)
	DIM arrMemberOf, strGroupDN
	DIM objGroup
	ON ERROR RESUME NEXT
	arrMemberOf = objUser.GetEx("memberOf")
 
	If Err.Number = E_ADS_PROPERTY_NOT_FOUND Then
    		exit sub
	End If
	ON ERROR GOTO 0
 
	For Each strGroupDN in arrMemberOf
    		Set objGroup = GetObject("LDAP://" & strGroupDN) 
    		objGroup.PutEx ADS_PROPERTY_DELETE, _
        		"member", Array(objUser.Get("distinguishedName"))
    		objGroup.SetInfo
	Next

	
end sub


private function getGroupDN(byval GroupName)
	DIM cmdGrp,cnGrp,rsGrp
	set cmdGrp=createobject("ADODB.Command")
	set cnGrp=createobject("ADODB.Connection")
	set rsGrp=createobject("ADODB.Recordset")
	
	cnGrp.open "Provider=ADsDSOObject;"
	
	cmdGrp.commandtext = "SELECT distinguishedName from 'LDAP://" & getDomainRoot() & _
			  "' WHERE objectCategory = 'Group' and sAMAccountName = '" & groupname & "'"
	cmdGrp.activeconnection = cnGrp
	
	set rsGrp = cmdGrp.execute

	if rsGrp.bof <> true and rsGrp.eof<>true then
		getgroupDN=rsGrp(0)
	else
		getgroupDN = ""
	end if
	cnGrp.close

end function

private sub addToGroup(userDN, groupDN)
	
	dim objGroup, objMember