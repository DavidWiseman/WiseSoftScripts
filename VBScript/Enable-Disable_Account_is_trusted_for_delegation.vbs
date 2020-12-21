const ADS_UF_ACCOUNT_TRUSTED = &H80000

Dim objUser
'<<<< Bind to the user object using the distinguished name >>>>
set objUser = getobject("LDAP://cn=test.3,cn=users,dc=wisesoft,dc=co,dc=uk")

intUAC = objUser.Get("userAccountControl")

'<<<<< Enable Account is trusted for delegation >>>>>
if  (intUAC AND ADS_UF_ACCOUNT_TRUSTED)=0 Then
	objUser.put "userAccountControl",  intUAC XOR ADS_UF_ACCOUNT_TRUSTED
	objUser.setinfo
end if

'<<<<< Disable Account is trusted for delegation >>>>>
if intUAC and ADS_UF_ACCOUNT_TRUSTED Then
	objUser.put "userAccountControl",  intUAC XOR ADS_UF_ACCOUNT_TRUSTED
	objUser.setinfo
end if