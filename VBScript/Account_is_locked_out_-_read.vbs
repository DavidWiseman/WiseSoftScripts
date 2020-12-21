DIM objUser
'<<<< Bind to the user object using the distinguished name >>>>
objUser = GetObject("LDAP://cn=test.3,cn=users,dc=wisesoft,dc=co,dc=uk")
'
if isAccountLocked(objUser) Then
	wscript.echo "The Account is locked out"
else
	wscript.echo "Account is not locked out"
end if

'<<<<< Function to check Account Lockout Status >>>>>
Function IsAccountLocked(byval objUser)
	on error resume next
	set objLockout = objUser.get("lockouttime")

	if err.number = -2147463155 then
		isAccountLocked = False
		exit Function
	end if
	on error goto 0
	
	if objLockout.lowpart = 0 And objLockout.highpart = 0 Then
		isAccountLocked = False
	Else
		isAccountLocked = True
	End If

End Function