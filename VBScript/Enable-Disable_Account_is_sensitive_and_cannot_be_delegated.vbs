const ADS_UF_ACCOUNT_SENSITIVE = &H100000

Dim objUser
'<<<< Bind to the user object using the distinguished name >>>>
set objUser = getobject("LDAP://cn=test.3,cn=users,dc=wisesoft,dc=co,dc=uk")

intUAC = objUser.Get("userAccountControl")

'<<<<< Enable Account is sensitive and cannot be delegated >>>>>
if  (intUAC AND ADS_UF_ACCOUNT_SENSITIVE)=0 Then
	objUser.put "userAccountControl",  intUAC XOR ADS_UF_ACCOUNT_SENSITIVE
	objUser.setinfo
end if

'<<<<< Disable Account is sensitive and cannot be delegated >>>>>
if intUAC and ADS_UF_ACCOUNT_SENSITIVE Then
	objUser.put "userAccountControl",  intUAC XOR ADS_UF_ACCOUNT_SENSITIVE
	objUser.setinfo
end if