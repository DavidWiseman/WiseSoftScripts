' Server Name | Server Name\Instance Name
strDBServerName = "." ' Local Server

Set objSQLServer = CreateObject("SQLDMO.SQLServer")
objSQLServer.LoginSecure = True
objSQLServer.Connect strDBServerName

WScript.Echo "SQL Major Version: " & objSQLServer.VersionMajor
WScript.Echo "SQL Minor Version: " & objSQLServer.VersionMinor
WScript.Echo "SQL Version String: " & objSQLServer.VersionString
