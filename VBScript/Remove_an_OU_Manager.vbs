Const ADS_PROPERTY_CLEAR = 1 
 
Set objContainer = GetObject _
  ("LDAP://ou=Students,dc=wisesoft,dc=co,dc=uk")

objContainer.PutEx ADS_PROPERTY_CLEAR, "managedBy", 0
objContainer.SetInfo
