Set objContainer = GetObject("LDAP://dc=wisesoft,dc=co,dc=uk")

objContainer.Delete "organizationalUnit", "ou=Staff"