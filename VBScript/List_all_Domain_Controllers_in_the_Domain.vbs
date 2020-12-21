Set cn = CreateObject("ADODB.Connection")
Set cmd= CreateObject("ADODB.Command")
cn.Provider = "ADsDSOObject;"
cn.open
cmd.ActiveConnection = cn

' Root DSE required to get the default configuration naming context to
' be used as the root of the seach
set objRootDSE = getobject("LDAP://RootDSE")
' Construct the LDAP query that will find all the domain controllers
' in the domain
ldapQuery = "<LDAP://" & objRootDSE.Get("ConfigurationNamingContext") & _
	">;((objectClass=nTDSDSA));ADsPath;subtree"

cmd.CommandText = ldapQuery
cmd.Properties("Page Size") = 1000
Set rs = cmd.Execute

do while rs.EOF <> True and rs.BOF <> True
	' Bind to the domain controller computer object
	' (This is the parent object of the result from the query)
	set objDC = getobject(getobject(rs(0)).Parent)

	wscript.echo objDC.dNSHostName
    	rs.MoveNext
Loop
	
cn.close