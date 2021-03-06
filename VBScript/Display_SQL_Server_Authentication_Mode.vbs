const SQLDMOSecurity_Integrated  = 1 
const SQLDMOSecurity_Mixed = 2 
const SQLDMOSecurity_Normal = 0
const SQLDMOSecurity_Unknown = 9 

strDBServerName = "." ' Local Server

Set objSQLServer = CreateObject("SQLDMO.SQLServer")

Select Case objSQLServer.ServerLoginMode(strDBServerName)
	Case SQLDMOSecurity_Integrated
      		WScript.Echo "Login Mode: Allow Windows Authentication only."
   	Case SQLDMOSecurity_Mixed
      		WScript.Echo "Login Mode: Allow Windows Authentication or SQL Server Authentication."
   	Case SQLDMOSecurity_Normal '?
      		WScript.Echo "Login Mode: Allow SQL Server Authentication only."
   	Case SQLDMOSecurity_Unknown '?
     		WScript.Echo "Login Mode: Security type unknown."
End Select