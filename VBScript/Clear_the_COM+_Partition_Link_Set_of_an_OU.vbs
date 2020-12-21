Const ADS_PROPERTY_CLEAR = 1 

Set objContainer = GetObject _
  ("LDAP://ou=Developers,dc=wisesoft,dc=co,dc=uk")
 
objContainer.PutEx ADS_PROPERTY_CLEAR, "msCOM-UserPartitionSetLink", 0
objContainer.SetInfo
