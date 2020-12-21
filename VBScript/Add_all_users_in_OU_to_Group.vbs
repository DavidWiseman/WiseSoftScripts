OPTION EXPLICIT

dim strFilter, strRoot, strScope, strGroupName
dim strNETBIOSDomain, strGroupDN
dim cmd, rs,cn, objGroup
dim objSystemInfo
' ********************* Setup *********************

' Default filter for all user accounts (ammend if required)
strFilter = "(&(objectCategory=person)(objectClass=user))"
' scope of search (default is subtree - search all child OUs)
strScope = "subtree"

' search root. e.g. ou=MyUsers,dc=wisesoft,dc=co,dc=uk
strRoot = "OU=Developers,dc=wisesoft,dc=co,dc=uk"

' Group to add
strGroupName = "Developers"

' *************************************************

SET objSystemInfo = CREATEOBJECT("ADSystemInfo") 
'Required for name translate
strNETBIOSDomain = objSystemInfo.DomainShortName

' Convert group name to distinguished name
strGroupDN =  GetDN(strNETBIOSDomain,strGroupName)

set objGroup = getobject("LDAP://" & strGroupDN)
set cmd = createobject("ADODB.Command")
set cn = createobject("ADODB.Connection")
set rs = createobject("ADODB.Recordset")

cn.open "Provider=ADsDSOObject;"
cmd.activeconnection = cn

cmd.commandtext = "<LDAP://" & strRoot & ">;" & strFilter & ";ADsPath,sAMAccountName;" & strScope

'**** Bypass 1000 record limitation ****
cmd.properties("page size")=1000

set rs = cmd.execute

' for each item returned by the Active Directory query
while rs.eof <> true and rs.bof <> true

	on error resume next
	' Add user to group
	objGroup.Add rs("ADsPath")

	' Error handling (user might already be a member of the group)
	if err.number = -2147019886 then
		wscript.echo "User '" & rs("sAMAccountName") & _
			"' is already a member of the '" & objGroup.sAMAccountName & "' group"
	elseif err.number <> 0 then
		wscript.echo "Error: " & err.number & " adding user '" & _
			 rs("sAMAccountName") & "' to the '" & objGroup.sAMAccountName & "' group"
	else
		wscript.echo "Added user '" & rs("sAMAccountName") & "' to the " & _
			objGroup.sAMAccountName & "' group"
	end if
	err.clear

	on error goto 0

	rs.movenext
wend

' Close ADO connection
cn.close

' Function to convert name into distinguished name format
Function GetDN(byval strDomain,strObject)
	' Use name translate to return the distinguished name
	' of a user from the NT UserName (sAMAccountName)
	' and the NETBIOS domain name.
	DIM objTrans

	Set objTrans = CreateObject("NameTranslate")
	objTrans.Init 1, strDomain
	objTrans.Set 3, strDomain & "\" & strObject
	GetDN = objTrans.Get(1) 

end function
