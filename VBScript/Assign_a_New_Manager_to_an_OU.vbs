Set objContainer = GetObject _
  ("LDAP://ou=Developers,dc=wisesoft,dc=co,dc=uk")
 
objContainer.Put "managedBy", "cn=david.wiseman,ou=Developers,dc=wisesoft,dc=co,dc=uk"
objContainer.SetInfo
