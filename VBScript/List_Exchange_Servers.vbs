set cn = createobject("ADODB.Connection")
set cmd = createobject("ADODB.Command")
set rs = createobject("ADODB.Recordset")

set objRoot = getobject("LDAP://RootDSE")
configurationNC = objRoot.Get("configurationnamingcontext")

cn.open "Provider=ADsDSOObject;"
cmd.activeconnection = cn

cmd.commandtext = "<LDAP://" & configurationNC & _
		  ">;(objectCategory=msExchExchangeServer);name;subtree"
set rs = cmd.execute

while rs.eof<>true and rs.bof<>true
	wscript.echo rs(0)
	rs.movenext
wend

cn.close