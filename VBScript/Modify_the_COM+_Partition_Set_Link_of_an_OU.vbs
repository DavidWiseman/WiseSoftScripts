Set objContainer = GetObject _
    ("LDAP://ou=Students,dc=wisesoft,dc=co,dc=uk")
 
objContainer.Put "msCOM-UserPartitionSetLink", _
    "cn=PartitionSet1,cn=ComPartitionSets,cn=System,dc=wisesoft,dc=co,dc=uk"
objContainer.SetInfo
