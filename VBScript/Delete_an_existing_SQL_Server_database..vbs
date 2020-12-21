strDBServerName = "."
strDBToDelete = "ScriptingGuysTestDB"

Set objSQLServer = CreateObject("SQLDMO.SQLServer")
objSQLServer.LoginSecure = True
objSQLServer.Connect strDBServerName

objSQLServer.KillDatabase strDBToDelete
