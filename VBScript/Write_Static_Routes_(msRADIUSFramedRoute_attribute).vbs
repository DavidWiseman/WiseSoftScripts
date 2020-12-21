OPTION EXPLICIT
Const ADS_PROPERTY_CLEAR = 1
Const ADS_PROPERTY_APPEND = 3 
Const ADS_PROPERTY_UPDATE = 2 

DIM objUser

'<<<< Bind to the user object using the distinguished name >>>>
set objUser = GetObject("LDAP://cn=user1,cn=users,dc=wisesoft,dc=org,dc=uk")

' Clear Assigned Static Routes
objUser.PutEx ADS_PROPERTY_CLEAR,"msRADIUSFramedRoute",NULL
' Optional - used by the user dialog in ADU&C
objUser.PutEx ADS_PROPERTY_CLEAR,"msRASSavedFramedRoute",NULL
objUser.SetInfo
wscript.echo "Cleared Assigned Static Routes"

' Update Assigned Static Routes
objUser.PutEx ADS_PROPERTY_UPDATE,"msRADIUSFramedRoute",Array("128.168.0.0/16 0.0.0.0 5","10.1.0.0/16 0.0.0.0 1")
' Optional - used by the user dialog in ADU&C
objUser.PutEx ADS_PROPERTY_UPDATE,"msRASSavedFramedRoute",Array("128.168.0.0/16 0.0.0.0 5","10.1.0.0/16 0.0.0.0 1")
objUser.SetInfo
wscript.echo "Updated Addigned Static Routes"

' Add a new Static Route
objUser.PutEx ADS_PROPERTY_APPEND,"msRADIUSFramedRoute",Array("192.168.0.0/16 0.0.0.0 2")
' Optional - used by the user dialog in ADU&C
objUser.PutEx ADS_PROPERTY_APPEND,"msRASSavedFramedRoute",Array("192.168.0.0/16 0.0.0.0 2")
objUser.SetInfo
wscript.echo "Appended a new Static Route"