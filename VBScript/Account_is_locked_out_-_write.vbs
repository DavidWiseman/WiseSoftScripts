DIM objUser
'<<<< Bind to the user object using the distinguished name >>>>
objUser = GetObject("LDAP://cn=test.3,cn=users,dc=wisesoft,dc=co,dc=uk")
'Accounts are locked by Active Directory - You can unlock accounts using a script.
if isAccountLocked(objUser) Then
	objuser.put "lockoutTime", 0
	objuser.setinfo ' Save Changes
	wscript.echo "Account unlocked."
else
	wscript.echo "Account was not locked out"
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