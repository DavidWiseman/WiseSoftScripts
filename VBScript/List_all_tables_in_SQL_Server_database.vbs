ServerName = inputbox("Please enter the name of the SQL Server:" & _
		vbcrlf & "(Connection is made using a trusted connection)")
if ServerName = "" then wscript.quit
DatabaseName = inputbox("Please enter the name of the database:")
if DatabaseName = "" then wscript.quit

set cn = createobject("ADODB.Connection")
set cmd = createobject("ADODB.Command")
set rs = createobject("ADODB.Recordset")

cn.open "Provider=SQLOLEDB.1;Data Source = " & ServerName & _
	";Integrated Security = SSPI;Initial Catalog=" & DatabaseName
cmd.activeconnection = cn

cmd.commandtext = "exec sp_help"

set rs = cmd.execute

while rs.eof <> true and rs.bof <> true
	if rs("object_type") = "user table" then
		display = display & rs("Name") & vbcrlf
	end if
	rs.movenext
wend

cn.close

wscript.echo display