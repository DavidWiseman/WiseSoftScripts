OPTION EXPLICIT
Const ADS_PROPERTY_CLEAR = 1

DIM strSearchFilter, strSearchRoot, objRootDSE
DIM cn,cmd,rs, strSearchScope
DIM objNewOU, strNewOU

' ********************************************************
' * Setup
' ********************************************************

' Specify the distinguished name of the new OU
strNewOU = "ou=Marketing,ou=All Users,dc=wisesoft,dc=org,dc=uk"

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

' loop through the search results
while rs.eof<> true and rs.bof<>true
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