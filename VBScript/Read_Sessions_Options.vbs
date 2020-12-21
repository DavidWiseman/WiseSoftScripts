Const TS_SESSION_DISCONNECT = 0
Const TS_SESSION_END = 1
Const TS_SESSION_ANY_CLIENT = 0
Const TS_SESSION_ORIGINATING = 1

Dim objUser
'<<<< Bind to the user object using the distinguished name >>>>
set objUser = getobject("LDAP://cn=test.3,cn=users,dc=wisesoft,dc=co,dc=uk")

'<<<< End a disconnected session >>>>
If objUser.MaxDisconnectionTime = 0 Then
	WScript.Echo "End a disconnected session: Never"
Else
	WScript.Echo "End a disconnected session: " & objUser.MaxDisconnectionTime & " Minute(s)"
End If

'<<<< Active session limit >>>>
If objUser.MaxConnectionTime = 0 Then
	WScript.Echo "Active session limit: Never"
Else
	WScript.Echo "Active session limit: " & objUser.MaxConnectionTime & " Minute(s)"
End If

'<<<< Idle session limit >>>>
If objUser.MaxIdleTime = 0 Then
	WScript.Echo "Idle session limit: Never"
Else
	WScript.Echo "Idle session limit: " & objUser.MaxIdleTime & " Minute(s)"
End If

'<<<< When a session limit is reached or a connection is broken >>>>
If objUser.BrokenConnectionAction = TS_SESSION_DISCONNECT Then
	WScript.Echo "When a session limit is reached or a connection is broken: Disconnect from session"
Else 'TS_SESSION_END
	WScript.Echo "When a session limit is reached or a connection is broken: End session"
End If

'<<<< Allow reconnection >>>>
If objUser.ReconnectionAction = TS_SESSION_ANY_CLIENT Then
	WScript.Echo "Allow Reconnection: From any client"
Else 
	'Will equal TS_SESSION_ORIGINATING
	WScript.Echo "Allow Reconnection: From originating client only"
End If	