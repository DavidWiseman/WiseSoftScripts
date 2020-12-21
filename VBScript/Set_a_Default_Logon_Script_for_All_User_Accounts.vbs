OPTION EXPLICIT
Const ADS_PROPERTY_CLEAR = 1

DIM strSearchFilter, strSearchRoot, objRootDSE
DIM cn,cmd,rs, strSearchScope
DIM objUser, strAttribute, strNewValue

' ********************************************************
' * Setup
' ********************************************************
' Specify the attribute you want to update
strAttribute = "scriptPath"
' Specify a new value for the attribute
strNewValue = "\\server\logonscripts\logon.bat"

' Specify a filter to use for the search
' This filter returns user accounts without a "logon script" attribute specified
strSearchFilter = "(&(objectCategory=Person)(objectClass=User)(!scriptPath=*))"

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

Set cn = CreateObject("ADODB.Connection")
Set cmd =   CreateObject("ADODB.Command")
cn.open "Provider=ADsDSOObject;"

Set cmd.ActiveConnection = cn

cmd.CommandText = "<LDAP://" & strSearchRoot & ">;" & strSearchFilter & ";ADsPath;" & strSearchScope
cmd.Properties("Page Size") = 1000

Set rs = cmd.Execute

' loop through the search results
while rs.eof<> true and rs.bof<>true
	' Bind to the user object (Using ADsPath returned from the search)
	set objUser = GetObject(rs(0))
	
	Err.Clear
	ON ERROR RESUME NEXT
	' Update the specified attribute
	wscript.echo "Updating user '" & rs(0) & "'"
	IF strNewValue = "" THEN
		objUser.PutEx ADS_PROPERTY_CLEAR, strAttribute, null
	ELSE
		objUser.Put strAttribute, strNewValue
	END IF
	objUser.SetInfo ' Commit Changes

	' Check if update succeeded
	IF err.Number = 0 Then
		wscript.echo "Succeeded"
	ELSE
		wscript.echo "Error: " & err.number
		wscript.echo err.description
		err.clear
	End If
	ON ERROR GOTO 0

	rs.movenext
wend

rs.close
cn.close

private function getDomainRoot
	' Bind to RootDSE - this object is used to 
	' get the default configuration naming context
	' e.g. dc=wisesoft,dc=co,dc=uk

	set objRootDSE = getobject("LDAP://RootDSE")
	getDomainRoot = objRootDSE.Get("DefaultNamingContext")
end function