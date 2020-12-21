OPTION EXPLICIT
Const E_ADS_PROPERTY_NOT_FOUND  = &h8000500D
DIM objUser,msRADIUSFramedIPAddress

'<<<< Bind to the user object using the distinguished name >>>>
set objUser = GetObject("LDAP://cn=user1,cn=users,dc=wisesoft,dc=org,dc=uk")

ON ERROR RESUME NEXT
msRADIUSFramedIPAddress= objUser.get("msRADIUSFramedIPAddress")

IF Err.Number = E_ADS_PROPERTY_NOT_FOUND then
	wscript.echo "Static IP Address Not Assigned"
	err.clear
else
	wscript.echo IntegerToIPAddress(msRADIUSFramedIPAddress)
End If

' Function to convert Integer value to IP Address.
Function IntegerToIPAddress(intIP)
	Const FourthOctet = 1
	Const ThirdOctet = 256
	Const SecondOctet = 65536
	Const FirstOctet = 16777216
	dim strIP,intFirstRemainder,intSecondRemainder,intThirdRemainder
	If sgn(intIP) = -1 Then
        	strIP =  (256 + (int(intIP/FirstOctet))) & "."
        	intFirstRemainder = intIP mod FirstOctet
        	strIP = strIP &  (256 + (int(intFirstRemainder/SecondOctet))) & "."
        	intSecondRemainder = intFirstRemainder mod SecondOctet
        	strIP = strIP & (256 + (int(intSecondRemainder/ThirdOctet))) & "."
       		intThirdRemainder = intSecondRemainder mod ThirdOctet
        	strIP = strIP & (256 + (int(intThirdRemainder/FourthOctet)))
    	Else
        	strIP = int(intIP/FirstOctet) & "."
        	intFirstRemainder = intIP mod FirstOctet
        	strIP = strIP & int(intFirstRemainder/SecondOctet) & "."
        	intSecondRemainder = intFirstRemainder mod SecondOctet
        	strIP = strIP & int(intSecondRemainder/ThirdOctet) & "."
        	intThirdRemainder = intSecondRemainder mod ThirdOctet
        	strIP = strIP & int(intThirdRemainder/FourthOctet)
    	End If
	IntegerToIPAddress = strIP
end function
