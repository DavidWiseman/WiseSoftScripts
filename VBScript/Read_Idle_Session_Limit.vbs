Dim objUser
'<<<< Bind to the user object using the distinguished name >>>>
set objUser = getobject("LDAP://cn=test.3,cn=users,dc=wisesoft,dc=co,dc=uk")

If objUser.MaxIdleTime = 0 Then
	WScript.Echo "Idle session limit: Never"
Else
	WScript.Echo "Idle session limit: " & objUser.MaxIdleTime & " Minute(s)"
End If