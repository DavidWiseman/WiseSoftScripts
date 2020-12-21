OPTION Explicit
DIM cn,cmd,rs
DIM objRoot
DIM strRoot, strFilter, strScope

' *******************************************************************
' * Setup
' *******************************************************************
' Specify OU of computers you want to shutdown
strRoot = "cn=computers,dc=wisesoft,dc=org,dc=uk"
' Default filter for computer objects
' You might want to use a different filter.  By operating system for example:
' (&(objectCategory=Computer)(operatingSystem=Windows XP*))
strFilter = "(objectCategory=Computer)"
' Search child organizational units.  Use "onelevel" to search only the specified OU.
strScope = "subtree"

' *******************************************************************

set cmd = createobject("ADODB.Command")
set cn = createobject("ADODB.Connection")
set rs = createobject("ADODB.Recordset")

cn.open "Provider=ADsDSOObject;"
cmd.activeconnection = cn

cmd.commandtext = "<LDAP://" & strRoot & ">;" & strFilter & ";" & _
		  "name;" & strScope
'**** Bypass 1000 record limitation ****
cmd.properties("page size")=1000

set rs = cmd.execute

while rs.eof <> true and rs.bof <> true
	wscript.echo "Shutting Down " & rs("name") & "..."
	ShutDownComputer(rs("name"))

	rs.movenext
wend

cn.close

' Subroutine to shutdown a computer
private sub ShutDownComputer(byval strComputer)
	dim strShutDown,objShell
	
	' -s = shutdown, -t 60 = 60 second timeout, -f = force programs to close
	strShutdown = "shutdown.exe -s -t 60 -f -m \\" & strComputer

    	set objShell = CreateObject("WScript.Shell") 

    	objShell.Run strShutdown, 0, false

end sub	