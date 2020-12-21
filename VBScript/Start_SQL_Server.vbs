strDBServerName = "."

Set objSQLServer = CreateObject("SQLDMO.SQLServer")
objSQLServer.LoginSecure = True
objSQLServer.Name = strDBServerName

objSQLServer.Start False
