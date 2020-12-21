Const ADS_PROPERTY_UPDATE = 2

Set objContainer = GetObject _
    ("LDAP://ou=Students,dc=wisesoft,dc=co,dc=uk")
 
objContainer.Put "street","Building 1" & vbCrLf & "Street XYZ"
objContainer.Put "l", "Sunderland"
objContainer.Put "st", "Tyne & Wear"
objContainer.Put "postalCode", "AA1 1AA"
objContainer.Put "c", "UK"
objContainer.PutEx ADS_PROPERTY_UPDATE, _
    "description", Array("Students")
objContainer.SetInfo

