Dim objUser
'<<<< Bind to the user object using the distinguished name >>>>
set objUser = getobject("LDAP://cn=test.3,cn=users,dc=wisesoft,dc=co,dc=uk")

' Enable
objUser.ConnectClientDrivesAtLogon = 1
objUser.setinfo

' Disable
objUser.ConnectClientDrivesAtLogon = 0
objUser.setinfo