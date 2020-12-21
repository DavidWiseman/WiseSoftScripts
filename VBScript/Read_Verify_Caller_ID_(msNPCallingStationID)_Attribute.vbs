Const E_ADS_PROPERTY_NOT_FOUND  = &h8000500D
DIM objUser,msNPCallingStationID
'<<<< Bind to the user object using the distinguished name >>>>
set objUser = GetObject("LDAP://cn=user1,cn=users,dc=wisesoft,dc=org,dc=uk")

' Error handling is required to check for null values
ON ERROR RESUME NEXT
msNPCallingStationID = objUser.get("msNPCallingStationID")
IF Err.Number = E_ADS_PROPERTY_NOT_FOUND then
	wscript.echo "Verify Caller ID: {Not Set}"
elseif blnMsNPAllowDialin = TRUE then
	wscript.echo "Verify Caller ID: " & msNPCallingStationID
else
	wscript.echo "Verify Caller ID: " & msNPCallingStationID
end if
