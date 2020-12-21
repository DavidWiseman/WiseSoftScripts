Const ADS_PROPERTY_CLEAR = 1

DIM objUser
'<<<< Bind to the user object using the distinguished name >>>>
set objUser = GetObject("LDAP://cn=user1,cn=users,dc=wisesoft,dc=org,dc=uk")

' Clear the Verify Caller ID
objUser.PutEx ADS_PROPERTY_CLEAR, "msNPCallingStationID", null
' Optional (Used by the ADU&C user interface)
objUser.PutEx ADS_PROPERTY_CLEAR, "msNPSavedCallingStationID", null
objUser.setinfo 'Save Changes
wscript.echo "Verify Caller ID attribute cleared"

' Write Verify Caller ID value
objUser.Put "msNPCallingStationID", "YourCallerID"
' Optional (Used by the ADU&C user interface)
objUser.Put "msNPSavedCallingStationID", "YourCallerID"
objUser.setinfo 'Save Changes
wscript.echo "Verify Caller ID attribute updated"