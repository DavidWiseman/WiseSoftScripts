strDBServerName = "."
strDBName = "Northwind"

Set objSQLServer = CreateObject("SQLDMO.SQLServer")
objSQLServer.LoginSecure = True
objSQLServer.Connect strDBServerName

Set colViews = objSQLServer.Databases(strDBName).Views
For Each objView In colViews
   WScript.Echo objView.Name
Next
