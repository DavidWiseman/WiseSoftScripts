dim objUser
'<<<< Bind to the user object using the distinguished name >>>>
set objUser = getobject("LDAP://cn=test.3,cn=users,dc=wisesoft,dc=co,dc=uk")

'<<<<< Read Settings >>>>>

WScript.echo "Terminal Services Profile Path : " & objUser.TerminalServicesProfilePath 
WScript.echo "Terminal Services Home Directory: " & objUser.TerminalServicesHomeDirectory
WScript.echo "Terminal Services Home Drive: " & objUser.TerminalServicesHomeDrive
WScript.echo "Allow Logon: " & objUser.AllowLogon

'<<<<< Write Settings >>>>>

objUser.TerminalServicesProfilePath = "\\server\tsprofiles\"
objUser.TerminalServicesHomeDirectory = "\\server\tshome\"
objUser.TerminalServicesHomeDrive = "H:"
objUser.AllowLogon = False

objUser.setinfo	