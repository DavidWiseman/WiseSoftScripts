Set objContainer = GetObject _
  ("LDAP://ou=Students,dc=wisesoft,dc=co,dc=uk")
 
Set objNtSecurityDescriptor = objContainer.Get("ntSecurityDescriptor")
 
WScript.Echo "Owner Tab"
WScript.Echo "Current owner of this item: " & objNtSecurityDescriptor.Owner
