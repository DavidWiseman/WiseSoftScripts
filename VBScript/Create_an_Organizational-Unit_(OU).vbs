Set objContainer = GetObject("LDAP://dc=wisesoft,dc=co,dc=uk")

Set objOU = objContainer.Create("organizationalUnit", "ou=Staff")
objOU.SetInfo