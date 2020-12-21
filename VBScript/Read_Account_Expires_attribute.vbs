'<<<< Bind to the user object using the distinguished name >>>>
set objUser = getobject("LDAP://cn=test.3,cn=users,dc=wisesoft,dc=co,dc=uk")

'<<<< Read Account Expiration Date >>>>>
' Error checking must be used as an error is thrown if the account expiration date is not set.
on error resume next 
dtmAccountExpiration = objUser.AccountExpirationDate 
If err.number = -2147467259 Or (datediff("d","01/01/1970",dtmAccountExpiration)<=0) Then 
	wscript.echo "No account expiration specified" 
Else 
   	wscript.echo objUser.AccountExpirationDate 
End If
on error goto 0 ' turn off resume next error handling