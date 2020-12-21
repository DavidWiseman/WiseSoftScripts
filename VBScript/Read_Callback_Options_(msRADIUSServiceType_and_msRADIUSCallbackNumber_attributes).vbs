Const E_ADS_PROPERTY_NOT_FOUND  = &h8000500D
DIM objUser,msRADIUSServiceType
'<<<< Bind to the user object using the distinguished name >>>>
set objUser = GetObject("LDAP://cn=user1,cn=users,dc=wisesoft,dc=org,dc=uk")

' A null value is used for "No Callback"
' Error handling is required
ON ERROR RESUME NEXT
msRADIUSServiceType= objUser.get("msRADIUSServiceType")

IF Err.Number = E_ADS_PROPERTY_NOT_FOUND then
	wscript.echo "No Callback"
	err.clear
elseif msRADIUSServiceType = 4 then
	msRADIUSCallbackNumber = objUser.Get("msRADIUSCallbackNumber")

	if msRADIUSCallbackNumber = "" then
		wscript.echo "Set by Caller"
	else
		wscript.echo "Always Callback to: " & msRADIUSCallbackNumber
	end if
else
	wscript.echo "Unknown Type:" & msRADIUSServiceType
end if
