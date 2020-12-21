On Error Resume Next
 
Set objContainer = GetObject _
   ("LDAP://ou=Students,dc=wisesoft,dc=co,dc=uk")
 
strManagedBy = objContainer.Get("managedBy")
 
If IsEmpty(strManagedBy) = TRUE Then
    WScript.Echo "No user account is assigned to manage " & _
        "this OU."
Else
    Set objUser = GetObject("LDAP://" & strManagedBy)
    WScript.Echo "Manager: " & objUser.streetAddress
    WScript.Echo "Office: " & _
      objUser.physicalDeliveryOfficeName  
    WScript.Echo "Street Address: " & strStreetAddress
    WScript.Echo "Locality: " & objUser.l
    WScript.Echo "State/province: " & objUser.st
    WScript.Echo "Country: " & objUser.c
    WScript.Echo "Telephone Number: " & objUser.telephoneNumber
    WScript.Echo "Fax Number: " & _
      objUser.facsimileTelephoneNumber
End If
