Const TS_SESSION_DISCONNECT = 0
Const TS_SESSION_END = 1

Dim objUser
'<<<< Bind to the user object using the distinguished name >>>>
set objUser = getobject("LDAP://cn=test.3,cn=users,dc=wisesoft,dc=co,dc=uk")

' Disconnect from session
objUser.BrokenConnectionAction = TS_SESSION_DISCONNECT '* Disconnect from session
objUser.setinfo

' End session
objUser.BrokenConnectionAction = TS_SESSION_END '* End session
objUser.setinfo