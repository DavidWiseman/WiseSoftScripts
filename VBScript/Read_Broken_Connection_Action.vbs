Const TS_SESSION_DISCONNECT = 0
Const TS_SESSION_END = 1

Dim objUser
'<<<< Bind to the user object using the distinguished name >>>>
set objUser = getobject("LDAP://cn=test.3,cn=users,dc=wisesoft,dc=co,dc=uk")

If objUser.BrokenConnectionAction = TS_SESSION_DISCONNECT Then
	WScript.Echo "When a session limit is reached or a connection is broken: Disconnect from session"
Else 'TS_SESSION_END
	WScript.Echo "When a session limit is reached or a connection is broken: End session"
End If