Dim objUser
'<<<< Bind to the user object using the distinguished name >>>>
set objUser = getobject("LDAP://cn=test.3,cn=users,dc=wisesoft,dc=co,dc=uk")

If objUser.MaxConnectionTime = 0 Then
	WScript.Echo "Active session limit: Never"
Else
	WScript.Echo "Active session limit: " & objUser.MaxConnectionTime & " Minute(s)"
End If