const ADS_UF_KERBEROS_PREAUTH = &H400000

'<<<< Bind to the user object using the distinguished name >>>>
set objUser = getobject("LDAP://cn=test.3,cn=users,dc=wisesoft,dc=co,dc=uk")

intUAC = objUser.Get("userAccountControl")

'<<<<< Enable Do not require Kerberos preauthentication >>>>>
if  (intUAC AND ADS_UF_KERBEROS_PREAUTH)=0 Then
	objUser.put "userAccountControl",  intUAC XOR ADS_UF_KERBEROS_PREAUTH
	objUser.setinfo
end if

'<<<<< Disable Do not require Kerberos preauthentication >>>>>
if intUAC and ADS_UF_KERBEROS_PREAUTH Then
	objUser.put "userAccountControl",  intUAC XOR ADS_UF_KERBEROS_PREAUTH
	objUser.setinfo
end if