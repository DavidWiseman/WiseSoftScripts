' Specify OU & Password
strContainer = "OU=Students,DC=wisesoft,DC=co,dc=uk"
strNewPassword = "p@sswORD"

' Variable to count number of password changes
intPasswordChanges = 0
' Bind to container
Set objContainer = GetObject("LDAP:// " & strContainer)
On Error Resume Next
' Filter for user objects
' Note: other objects also inherit from user class. 
' e.g. Computer objects
objContainer.Filter = Array("User")

For each objUser in objContainer
	' Check that object is a user account
	if left(objUser.objectCategory,9) = "CN=Person" then
		' Set password
		objUser.SetPassword (strNewPassword)
		' Display message if error changing password
   		If Err.Number <> 0 Then
      			Wscript.Echo "Unable to change password for user: " & _
				objUser.sAMAccountName
   			Err.Clear
		else
			' Keep count of the number of password changes
			intPasswordChanges = intPasswordChanges + 1
   		End If
	end if
Next
' Display message
wscript.echo "Finished" & vbcrlf & intPasswordChanges & _
	" Passwords changed"