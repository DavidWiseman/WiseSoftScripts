option explicit

dim cn
dim cmd
dim rs
dim objRoot
dim strFilter
dim namingContext 
dim query
dim objUser

' Get the default naming context to be used in query later
' e.g. dc=wisesoft,dc=co,dc=uk
set objRoot = getobject("LDAP://RootDSE")
namingContext = objRoot.get("defaultNamingContext")

' Filter all users with a lockout time (*) that is not equal to zero
' - some will be locked out, others will have been unlocked automatically by the
' domain lockout duration policy. The lockout time will be set to zero after a 
' succesful logon for these users and also after execution of this script.
' Although it isn't necesary to unlock these users - you would have to compare 
' the lockouttime with the current time and the lockout duration policy for the 
' domain (requires conversion from 64bit numbers)
strFilter = "(&(objectCategory=person)(objectClass=user)(lockouttime=*)(!lockoutTime=0))"

' Query string will use naming context as the base, the filter above to filter locked out users*, 
' will return the adspath of the user (we can create a user object from this) and will search all 
' OU's within the context specified (the domain)
query = "<LDAP://" & namingContext & ">;" & strFilter & ";adspath;subtree"

set cmd = createobject("ADODB.Command")
set cn =createobject("ADODB.Connection")
set rs = createobject("ADODB.Recordset")
cn.open "Provider=ADsDSOObject;"

cmd.activeconnection = cn
cmd.commandtext = query

' Bypass 1000 record limitation ****
cmd.properties("page size")=1000

' Here is where the query is actually executed
set rs = cmd.execute

' Enumerate all rows returned from the query and set users lockout time to 0 (unlock account)
while rs.eof <> true and rs.bof <> true

	' Bind to user using ADsPath obtained from query
	set objUser = getobject(rs(0))
	
	' Write lockout time value to unlock account
	objuser.put "lockoutTime", 0
	objUser.setinfo

	rs.movenext
wend

wscript.echo "Unlocked " & rs.recordcount & " accounts"