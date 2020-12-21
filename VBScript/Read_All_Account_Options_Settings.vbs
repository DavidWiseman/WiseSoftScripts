const ADS_UF_PASSWD_CANT_CHANGE = &H40
const ADS_UF_DONT_EXPIRE_PASSWD = &H10000
const ADS_UF_PASSWORD_EXPIRED = &H800000
const ADS_UF_ACCOUNTDISABLE = &H02
const ADS_UF_ENCRYPTED_TEXT_PASSWD = &H80
const ADS_UF_SMARTCARD_REQUIRED = &h40000
const ADS_UF_ACCOUNT_TRUSTED = &H80000
const ADS_UF_ACCOUNT_SENSITIVE = &H100000
const ADS_UF_DES_ENCRYPTION = &H200000
const ADS_UF_KERBEROS_PREAUTH = &H400000

Dim objUser
'<<<< Bind to the user object using the distinguished name >>>>
set objUser = getobject("LDAP://cn=test.3,cn=users,dc=wisesoft,dc=co,dc=uk")

intUAC = objUser.Get("userAccountControl")
wscript.echo intUAC

if objUser.Get("pwdLastSet").HighPart = 0 then
	message = message & "User must change password at next logon: TRUE" & vbcrlf
else
	message = message & "User must change password at next logon: FALSE" & vbcrlf
end if

if getUserCannotChangePWD(objUser) = TRUE then
	message = message & "User cannot change password: TRUE" & vbcrlf
else
	message = message & "User cannot change password: FALSE" & vbcrlf
end if
if intUAC and ADS_UF_DONT_EXPIRE_PASSWD then
	message = message & "Password never expires: TRUE" & vbcrlf
else
	message = message & "Password never expires: FALSE" & vbcrlf
end if

if intUAC and ADS_UF_ENCRYPTED_TEXT_PASSWD then
	message = message & "Store password using reversible encryption: TRUE" & vbcrlf
else
	message = message & "Store password using reversible encryption: FALSE" & vbcrlf
end if

if intUAC and ADS_UF_ACCOUNTDISABLE then
	message = message & "Account Disabled: TRUE" & vbcrlf
else
	message = message & "Account Disabled: FALSE" & vbcrlf
end if

if intUAC and ADS_UF_SMARTCARD_REQUIRED then
	message = message & "Smart Card is required for interactive logon: TRUE" & vbcrlf
else
	message = message & "Smart Card is required for interactive logon: FALSE" & vbcrlf
end if

if intUAC and ADS_UF_ACCOUNT_TRUSTED then
	message = message & "Account is trusted for delegation: TRUE" & vbcrlf
else
	message = message & "Account is trusted for delegation: FALSE" & vbcrlf
end if

if intUAC and ADS_UF_ACCOUNT_SENSITIVE then
	message = message & "Account is sensitive and cannot be delegated: TRUE" & vbcrlf
else
	message = message & "Account is sensitive and cannot be delegated: FALSE" & vbcrlf
end if

if intUAC and ADS_UF_DES_ENCRYPTION then
	message = message & "Use DES encryption types for this account: TRUE" & vbcrlf
else
	message = message & "Use DES encryption types for this account: FALSE" & vbcrlf
end if

if intUAC and ADS_UF_KERBEROS_PREAUTH then
	message = message & "Do not require Kerberos preauthentication: TRUE" & vbcrlf
else
	message = message & "Do not require Kerberos preauthentication: FALSE" & vbcrlf	
end if

if intUAC and ADS_UF_PASSWD_CANT_CHANGE then
	message = message & "User cannot change password: TRUE" & vbcrlf
else
	message = message & "User cannot change password: FALSE" & vbcrlf
end if

wscript.echo message


'<<<<< Function to return if user cannot change password has been set >>>>>>
function getUserCannotChangePWD(byval objUser)

	Const CHANGE_PASSWORD_GUID = "{AB721A53-1E2F-11D0-9819-00AA0040529B}"
	Const ADS_ACETYPE_ACCESS_ALLOWED_OBJECT = &H5
	' Bind to the user security objects.
	Set objSecDescriptor = objUser.Get("ntSecurityDescriptor")
	Set objDACL = objSecDescriptor.discretionaryAcl

	For Each objACE In objDACL
  		If UCase(objACE.objectType) = UCase(CHANGE_PASSWORD_GUID) Then
    			If UCase(objACE.Trustee) = "NT AUTHORITY\SELF" Then
      				If objACE.AceType = ADS_ACETYPE_ACCESS_ALLOWED_OBJECT Then
        				getUserCannotChangePWD=False
					exit function
      				End If
    			End If
    			If UCase(objACE.Trustee) = "EVERYONE" Then
      				If objACE.AceType = ADS_ACETYPE_ACCESS_ALLOWED_OBJECT Then
        				getUserCannotChangePWD=False
					exit Function
      				End If
    			End If
  		End If
	Next

    	getUserCannotChangePWD=True
end Function