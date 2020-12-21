strDBServerName = "."

Set objSQLServer = CreateObject("SQLDMO.SQLServer")
objSQLServer.LoginSecure = True
objSQLServer.Connect strDBServerName

Set colDatabases = objSQLServer.Databases

For Each objDatabase In colDatabases
   WScript.Echo objDatabase.Name
Next
