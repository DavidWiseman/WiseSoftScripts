' Local Computer
strDBServerName = "."
strDBName = "Northwind" ' Change database as required

Set objSQLServer = CreateObject("SQLDMO.SQLServer")
' Integrated Security
objSQLServer.LoginSecure = True
objSQLServer.Connect strDBServerName

Set objDB = objSQLServer.Databases(strDBName)
Set colStoredProcedures = objDB.StoredProcedures
' For each stored procedure...
For Each objStoredProcedure In colStoredProcedures
	WScript.Echo "SP Name: " & objStoredProcedure.Name
Next

