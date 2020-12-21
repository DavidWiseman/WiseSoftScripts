DB_SIZE_IN_MEGABYTES = 5
strDBServerName = "."

Set objSQLServer = CreateObject("SQLDMO.SQLServer")
objSQLServer.LoginSecure = True
objSQLServer.Connect strDBServerName

Set objDB = CreateObject("SQLDMO.Database")
Set objFG = CreateObject("SQLDMO.Filegroup")
Set objDBFile = CreateObject("SQLDMO.DBFile")
Set objLogFile = CreateObject("SQLDMO.LogFile")

objDB.Name = "ScriptingGuysTestDB"
objDBFile.Name = "ScriptingGuysTestDB_Data"
objDBFile.PhysicalName = "C:\Program Files\Microsoft SQL Server\MSSQL\data\ScriptingGuysTestDB_Data.MDF"
objDBFile.Size = DB_SIZE_IN_MEGABYTES

objDB.FileGroups("PRIMARY").DBFiles.Add(objDBFile)

objSQLServer.Databases.Add(objDB)
