const ADS_UF_SMARTCARD_REQUIRED = &h40000

Dim objUser
'<<<< Bind to the user object using the distinguished name >>>>
set objUser = getobject("LDAP://cn=test.3,cn=users,dc=wisesoft,dc=co,dc=uk")

intUAC = objUser.Get("userAccountControl")

'<<<<< Enable Smart Card is required for interactive logon >>>>>
if  (intUAC AND ADS_UF_SMARTCARD_REQUIRED)=0 Then
	objUser.put "userAccountControl",  intUAC XOR ADS_UF_SMARTCARD_REQUIRED
	objUser.setinfo
end if

'<<<<< Disable Smart Card is required for interactive logon >>>>>
if intUAC and ADS_UF_SMARTCARD_REQUIRED Then
	objUser.put "userAccountControl",  intUAC XOR ADS_UF_SMARTCARD_REQUIRED
	objUser.setinfo
end if