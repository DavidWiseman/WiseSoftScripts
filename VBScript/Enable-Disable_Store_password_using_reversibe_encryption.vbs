const ADS_UF_ENCRYPTED_TEXT_PASSWD = &H80

Dim objUser
'<<<< Bind to the user object using the distinguished name >>>>
set objUser = getobject("LDAP://cn=test.3,cn=users,dc=wisesoft,dc=co,dc=uk")

intUAC = objUser.Get("userAccountControl")

'<<<<< Enable Store password using reversible encryption >>>>>
if  (intUAC AND ADS_UF_ENCRYPTED_TEXT_PASSWD)=0 Then
	objUser.put "userAccountControl",  intUAC XOR ADS_UF_ENCRYPTED_TEXT_PASSWD
	objUser.setinfo
end if


'<<<<< Disable Store password using reversible encryption >>>>>
if intUAC and ADS_UF_ENCRYPTED_TEXT_PASSWD Then
	objUser.put "userAccountControl",  intUAC XOR ADS_UF_ENCRYPTED_TEXT_PASSWD
	objUser.setinfo
end if