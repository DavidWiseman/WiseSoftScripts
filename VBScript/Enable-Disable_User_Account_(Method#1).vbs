const ADS_UF_ACCOUNTDISABLE = &H02

Dim objUser
'<<<< Bind to the user object using the distinguished name >>>>
set objUser = getobject("LDAP://cn=test.3,cn=users,dc=wisesoft,dc=co,dc=uk")

intUAC = objUser.Get("userAccountControl")

'<<<<< Disable Account >>>>>
if  (intUAC AND ADS_UF_ACCOUNTDISABLE)=0 Then
	objUser.put "userAccountControl",  intUAC XOR ADS_UF_ACCOUNTDISABLE
	objUser.setinfo
end if

'<<<<< Enable Account >>>>>
if intUAC and ADS_UF_ACCOUNTDISABLE Then
	objUser.put "userAccountControl",  intUAC XOR ADS_UF_ACCOUNTDISABLE
	objUser.setinfo
end if