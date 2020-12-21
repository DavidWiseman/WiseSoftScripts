const ADS_UF_DES_ENCRYPTION = &H200000

'<<<< Bind to the user object using the distinguished name >>>>
set objUser = getobject("LDAP://cn=test.3,cn=users,dc=wisesoft,dc=co,dc=uk")

intUAC = objUser.Get("userAccountControl")

'<<<<< Enable Use DES encryption types for this account >>>>>
if  (intUAC AND ADS_UF_DES_ENCRYPTION)=0 Then
	objUser.put "userAccountControl",  intUAC XOR ADS_UF_DES_ENCRYPTION
	objUser.setinfo
end if

'<<<<< Disable Use DES encryption types for this account >>>>>
if intUAC and ADS_UF_DES_ENCRYPTION Then
	objUser.put "userAccountControl",  intUAC XOR ADS_UF_DES_ENCRYPTION
	objUser.setinfo
end if