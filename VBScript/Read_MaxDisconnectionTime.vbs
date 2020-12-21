Dim objUser
'<<<< Bind to the user object using the distinguished name >>>>
set objUser = getobject("LDAP://cn=test.3,cn=users,dc=wisesoft,dc=co,dc=uk")

'<<<< End a disconnected session >>>>
If objUser.MaxDisconnectionTime = 0 Then
	WScript.Echo "End a disconnected session: Never"
Else
	WScript.Echo "End a disconnected session: " & objUser.MaxDisconnectionTime & " Minute(s)"
End If