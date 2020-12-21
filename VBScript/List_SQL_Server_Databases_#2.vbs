set cn = createobject("ADODB.Connection")
set cmd = createobject("ADODB.Command")
set rs = createobject("ADODB.RecordSet")

'**** Change the connection string as required
cn.open "Provider=SQLOLEDB.1;Data Source=(local);Integrated Security=SSPI;Initial Catalog=Master"
cmd.activeconnection = cn

cmd.commandtext = "exec sp_helpdb"
set rs = cmd.execute

display = "Size" & vbtab & vbtab & "Name" & vbcrlf
while rs.eof <> true and rs.bof <> true
	display = display & trim(rs("db_size")) & vbtab & vbtab & rs("name") & vbcrlf
	rs.movenext
wend

cn.close

wscript.echo display