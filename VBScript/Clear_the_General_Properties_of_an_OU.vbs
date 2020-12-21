Const ADS_PROPERTY_CLEAR = 1 

Set objContainer = GetObject _
  ("LDAP://ou=Students,dc=wisesoft,dc=co,dc=uk")
 
objContainer.PutEx ADS_PROPERTY_CLEAR, "description", 0
objContainer.PutEx ADS_PROPERTY_CLEAR, "street", 0
objContainer.PutEx ADS_PROPERTY_CLEAR, "l", 0
objContainer.PutEx ADS_PROPERTY_CLEAR, "st", 0
objContainer.PutEx ADS_PROPERTY_CLEAR, "postalCode", 0
objContainer.PutEx ADS_PROPERTY_CLEAR, "c", 0
objContainer.SetInfo
