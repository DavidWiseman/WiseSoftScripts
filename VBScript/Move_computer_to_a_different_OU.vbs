set objOU = getobject("LDAP://ou=newOU,dc=wisesoft,dc=co,dc=uk")
objOU.MoveHere "LDAP://cn=computer1,cn=computers,dc=wisesoft,dc=co,dc=uk",vbNullString