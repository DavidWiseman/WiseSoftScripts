Const ADS_PROPERTY_CLEAR = 1 
 
Set objContainer = GetObject _
    ("LDAP://ou=students,dc=wisesoft,dc=co,dc=uk")

objContainer.PutEx ADS_PROPERTY_CLEAR, "gPLink", 0
objContainer.PutEx ADS_PROPERTY_CLEAR, "gPOptions", 0
objContainer.SetInfo
