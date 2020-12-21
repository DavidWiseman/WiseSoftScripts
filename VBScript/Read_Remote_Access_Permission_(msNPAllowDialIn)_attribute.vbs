Const E_ADS_PROPERTY_NOT_FOUND  = &h8000500D
DIM objUser,blnMsNPAllowDialin
'<<<< Bind to the user object using the distinguished name >>>>
set objUser = GetObject("LDAP://cn=user1,cn=users,dc=wisesoft,dc=org,dc=uk")

' A null value is used for "control access through remote access policy"
' this will cause an error when trying to read the msNPAllowDialIn attribute.
' Error handling is required
ON ERROR RESUME NEXT
blnMsNPAllowDialin = objUser.get("msNPAllowDialIn")
IF Err.Number = E_ADS_PROPERTY_NOT_FOUND then
	wscript.echo "Control access through remote access policy"
elseif blnMsNPAllowDialin = TRUE then
	wscript.echo "Allow Access"
else
	wscript.echo "Deny Access"
end if
