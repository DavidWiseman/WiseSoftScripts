Dim objUser
'<<<< Bind to the user object using the distinguished name >>>>
set objUser = getobject("LDAP://cn=test.3,cn=users,dc=wisesoft,dc=co,dc=uk")

'<<<< Read Account Expiration Date >>>>>
on error resume next
dtmAccountExpiration = objUser.AccountExpirationDate 
If err.number = -2147467259 Or (datediff("d","01/01/1970",dtmAccountExpiration)<=0) Then 
	wscript.echo "No account expiration specified" 
Else 
   	wscript.echo objUser.AccountExpirationDate 
End If
on error goto 0

'<<<< Write Account Expiration Date >>>>>
objUser.AccountExpirationDate = "09/09/2005"
objUser.setinfo

'<<<< Don't expire account >>>>>
objUser.AccountExpirationDate = 0
objUser.setinfo
