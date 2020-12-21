' Local Computer
strDBServerName = "."
strDBName = "Northwind"

Set objSQLServer = CreateObject("SQLDMO.SQLServer")
' Integrated Security
objSQLServer.LoginSecure = True
objSQLServer.Connect strDBServerName

Set colViews = objSQLServer.Databases(strDBName).Views
' List Views
For Each objView In colViews
	WScript.Echo objView.Name
Next
