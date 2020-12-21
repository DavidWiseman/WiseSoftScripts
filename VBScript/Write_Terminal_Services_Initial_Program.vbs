Dim objUser
'<<<< Bind to the user object using the distinguished name >>>>
set objUser = getobject("LDAP://cn=test.3,cn=users,dc=wisesoft,dc=co,dc=uk")

objUser.TerminalServicesInitialProgram = "C:\test.exe"
objUser.setinfo