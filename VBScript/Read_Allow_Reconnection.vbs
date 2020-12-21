Const TS_SESSION_ANY_CLIENT = 0
Const TS_SESSION_ORIGINATING = 1

Dim objUser
'<<<< Bind to the user object using the distinguished name >>>>
set objUser = getobject("LDAP://cn=test.3,cn=users,dc=wisesoft,dc=co,dc=uk")

If objUser.ReconnectionAction = TS_SESSION_ANY_CLIENT Then
	WScript.Echo "Allow Reconnection: From any client"
Else 
	'Will equal TS_SESSION_ORIGINATING
	WScript.Echo "Allow Reconnection: From originating client only"
End If