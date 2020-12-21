'**** Change this constant to display system databases
const DisplaySystemDatabases = false

set objSQLServer = createobject("SQLDMO.SQLServer")
objSQLServer.LoginSecure = True

'**** Change this connection as required
objSQLServer.Connect "(local)"


For i = 1 To objSQLServer.Databases.Count
	'***** Test to see if the database is a system database
	if (not objSQLServer.Databases(i).SystemObject) or DisplaySystemDatabases then
		wscript.echo objSQLServer.Databases(i).Name
	end if
Next