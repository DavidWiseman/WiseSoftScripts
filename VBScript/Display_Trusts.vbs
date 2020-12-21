set objRoot = getobject("LDAP://RootDSE")
defaultNC = objRoot.get("defaultNamingContext")

set cn = createobject("ADODB.Connection")
set cmd = createobject("ADODB.Command")
set rs = createobject("ADODB.Recordset")

cn.open "Provider=ADsDSOObject;"
cmd.activeconnection =cn

cmd.commandtext = "SELECT trustPartner,trustDirection, TrustType, flatName FROM 'LDAP://CN=System," & _
		  DefaultNC & "' WHERE objectclass = 'trusteddomain'"

set rs = cmd.execute

while rs.eof <> true and rs.bof <> true
	select case rs("trustDirection")
		case 0
			TrustDirection = "Disabled"
		case 1
			TrustDirection = "Inbound trust"
		case 2
			TrustDirection = "Outbound trust"
		case 3
			TrustDirection = "Two-way trust"
	end select
	select case rs("trustType")
		case 1
			TrustType = "Downlevel Trust"
		case 2
			TrustType = "Windows 2000 (Uplevel) Trust"
		case 3
			TrustType = "MIT"
		case 4
			TrustType = "DCE"
	end select
	wscript.echo "DNS DomainName: " & rs("trustPartner") & vbnewline & _
		     "NetBIOS Name: " & rs("flatName") & vbnewline & _
		     "Trust Type: " & TrustType & vbnewline & _
		     "Trust Direction: " & TrustDirection
	rs.movenext
wend


cn.close