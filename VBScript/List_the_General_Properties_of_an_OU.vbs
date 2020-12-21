On Error Resume Next

Set objContainer = GetObject _
  ("LDAP://ou=Students,dc=wisesoft,dc=co,dc=uk")
 
For Each strValue in objContainer.description
  WScript.Echo "Description: " & strValue
Next
 
Wscript.Echo "Street Address: " & strStreetAddress
Wscript.Echo "Locality: " & 
Wscript.Echo "State/porvince: " & objContainer.st
Wscript.Echo "Postal Code: " & objContainer.postalCode
Wscript.Echo "Country: " & objContainer.c
