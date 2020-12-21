Const ADS_PROPERTY_CLEAR = 1

DIM objUser
'<<<< Bind to the user object using the distinguished name >>>>
set objUser = GetObject("LDAP://cn=user1,cn=users,dc=wisesoft,dc=org,dc=uk")

' Set "No Callback"
objUser.PutEx ADS_PROPERTY_CLEAR, "msRADIUSServiceType", null
objUser.PutEx ADS_PROPERTY_CLEAR, "msRADIUSCallbackNumber", null
objUser.SetInfo
wscript.echo "Updated: No Callback"

' Set By Caller
objUser.Put "msRADIUSServiceType", 4
objUser.PutEx ADS_PROPERTY_CLEAR, "msRADIUSCallbackNumber", null
objUser.SetInfo
wscript.echo "Updated: Set By Caller"

' Always Callback To
objUser.Put "msRADIUSServiceType", 4
objUser.Put "msRADIUSCallbackNumber", "MyCallBackNumber"
' Optional - used by the user dialog in ADU&C
objUser.Put "msRASSavedCallbackNumber", "MyCallBackNumber"
objUser.SetInfo
wscript.echo "Updated: Always callback To..."
