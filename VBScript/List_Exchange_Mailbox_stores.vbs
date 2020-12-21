Set objRootDSE = CreateObject("LDAP://RootDSE")
configurationNamingContext = objRootDSE.get("configurationNamingContext")
   
Set cn = CreateObject("ADODB.Connection")
Set cmd = CreateObject("ADODB.Command")
Set rs = CreateObject("ADODB.Recordset")

cn.Open "Provider=ADsDSOObject;"
        
query = "<LDAP://" & configurationNamingContext & ">;(objectCategory=msExchPrivateMDB);name,cn,distinguishedName;subtree"

cmd.ActiveConnection = cn
cmd.CommandText = query
Set rs = cmd.Execute

While rs.EOF <> True And rs.BOF <> True
	wscript.echo rs.Fields("distinguishedname").Value
	rs.MoveNext
Wend
   
Set rs = Nothing
Set cmd = Nothing
Set cn = Nothing