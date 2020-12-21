set objOU = getobject("LDAP://ou=newOU,dc=wisesoft,dc=co,dc=uk")
objOU.MoveHere "LDAP://cn=user1,cn=users,dc=wisesoft,dc=co,dc=uk",vbNullString