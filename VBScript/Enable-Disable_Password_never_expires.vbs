const ADS_UF_DONT_EXPIRE_PASSWD = &H10000

dim objUser
'<<<< Bind to the user object using the distinguished name >>>>
set objUser = getobject("LDAP://cn=test.3,cn=users,dc=wisesoft,dc=co,dc=uk")

intUAC = objUser.Get("userAccountControl")

'<<<<< Enable Password never expires >>>>>
if  (intUAC AND ADS_UF_DONT_EXPIRE_PASSWD)=0 Then
	objUser.put "userAccountControl",  intUAC XOR ADS_UF_DONT_EXPIRE_PASSWD
	objUser.setinfo
end if

'<<<<< Disable Password never expires >>>>>
if intUAC and ADS_UF_DONT_EXPIRE_PASSWD Then
	objUser.put "userAccountControl",  intUAC XOR ADS_UF_DONT_EXPIRE_PASSWD
	objUser.setinfo
end if