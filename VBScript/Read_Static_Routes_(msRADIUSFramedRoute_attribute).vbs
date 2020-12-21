OPTION EXPLICIT
Const E_ADS_PROPERTY_NOT_FOUND  = &h8000500D
DIM objUser,arrMsRADIUSFramedRoute

'<<<< Bind to the user object using the distinguished name >>>>
set objUser = GetObject("LDAP://cn=user1,cn=users,dc=wisesoft,dc=org,dc=uk")

ON ERROR RESUME NEXT
arrMsRADIUSFramedRoute= objUser.GetEx("msRADIUSFramedRoute")
If Err.Number = E_ADS_PROPERTY_NOT_FOUND Then
	WScript.Echo "No static Routes Applied."
	Err.Clear
Else
	Dim strValue
    	' CIDR 0.0.0.0 Metric
    	For Each strValue in arrMsRADIUSFramedRoute
        	WScript.echo strValue
    	Next
End If
