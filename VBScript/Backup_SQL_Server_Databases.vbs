serverName = inputbox("Please enter the name of the SQL Server:" & vbcrlf & _
		"(connection will be made using a trusted connection)")
if serverName = "" then wscript.quit

backupDirectory = inputbox("Please enter the backup directory:")
if backupdirectory = "" then wscript.quit
if not mid(backupDirectory,len(backupDirectory),1)="\" then
	backupDirectory = backupDirectory & "\"
end if


set cn = createobject("ADODB.Connection")
set cmd = createobject("ADODB.Command")
set rs = createobject("ADODB.RecordSet")

cn.open "Provider=SQLOLEDB.1;Data Source=" & serverName & ";Integrated Security=SSPI;Initial Catalog=Master"
cmd.activeconnection = cn

cmd.commandtext = "exec sp_helpdb"
set rs = cmd.execute

while rs.eof <> true and rs.bof <> true
	if not rs("name") = "tempdb" then
		backupDatabase rs("name")
	end if
	rs.movenext
wend

cn.close

wscript.echo "Backup Complete"


sub backupDatabase(byval databaseName)
	
	fileName = backupDirectory & replace(replace(now,"/","_"),":","_") & "_" & databaseName & ".BKF" 

	set cmdbackup = createobject("ADODB.Command")
	cmdbackup.activeconnection = cn
	cmdbackup.commandtext = "backup database " & databaseName & " to disk='" & fileName & "'"
	cmdbackup.execute

end sub