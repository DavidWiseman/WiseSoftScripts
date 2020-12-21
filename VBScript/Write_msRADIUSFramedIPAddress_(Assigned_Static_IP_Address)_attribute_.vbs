OPTION EXPLICIT
Const ADS_PROPERTY_CLEAR = 1

DIM objUser
'<<<< Bind to the user object using the distinguished name >>>>
set objUser = GetObject("LDAP://cn=user1,cn=users,dc=wisesoft,dc=org,dc=uk")

' Clear Attribute (Static IP Address not Assigned)
objUser.PutEx ADS_PROPERTY_CLEAR,"msRADIUSFramedIPAddress", NULL
' Optional (Used by the user dialog in ADU&C)
objUser.PutEx ADS_PROPERTY_CLEAR,"msRASSavedFramedIPAddress", NULL
objUser.SetInfo
wscript.echo "Attribute Cleared - Static IP Address not assigned"

objUser.Put "msRADIUSFramedIPAddress", ipAddressToInteger( "192.168.1.10")
' Optional (Used by the user dialog in ADU&C)
objUser.Put "msRASSavedFramedIPAddress", ipAddressToInteger( "192.168.1.10")
objUser.SetInfo
wscript.echo "Attribute Updated"

' Function to convert an IP Address to an integer value
Function ipAddressToInteger(ByVal ipAddress)
	Dim octets,strBin,i, value
	octets = SPLIT(ipAddress,".")
        strBin = ""
	' Convert each octet from decimal values to binary
	' Append the binary values in order to strBin variable
	' e.g. Convert 192.168.1.10 to 11000000 10101000 00000001 00001010
        If UBOUND(octets) = 3 Then
		For i = 0 To 3
                	If IsNumeric(octets(i)) Then
				If octets(i) <= 255 Then
                        		strBin = strBin & RIGHT("00000000" & DecToBin(octets(i)), 8)
				Else
                        		err.raise "Invalid IP Address"
				End If
                	Else
				err.raise vbObjectError + 1,"IPAddressToInteger","Invalid IP Address"
                	End If
		Next
        Else
		err.raise vbObjectError + 1,"IPAddressToInteger","Invalid IP Address"
        End If
	' Convert binary value back to decimal
        value = BinToDec(strBin)
	' Convert large numbers to negative
	If value > 2147483647 Then 
		value = value - 4294967296
        End If
        ipAddressToInteger = CLNG(value)

End Function

' Function to convert a decimal number to binary
Function DecToBin(ByVal d)
	DIM value
	value = ""
	Do While d > 0
		If d Mod 2 > 0 Then
			value = "1" & value
		Else
			value = "0" & value
		End If
		d = Int(d / 2)
	Loop
	DecToBin = value
End Function 

' Function to convert a binary number to decimal
Function BinToDec(strBin)
  	dim value, i, strDigit

  	value = 0
  	for i = len(strBin) to 1 step -1
    		strDigit = mid(strBin, i, 1)
    		select case strDigit
      		case "0"
        		' do nothing
      		case "1"
        		value = value + (2 ^ (len(strBin)-i))
      		case else
			err.raise vbObjectError + 1,"BinToDec","Invalid Binary Digit"
    		end select
 	next

  	BinToDec = value
End Function
