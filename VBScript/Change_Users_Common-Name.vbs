DIM objOU
set objOU = getobject("LDAP://cn=users,dc=wisesoft,dc=ac,dc=uk")
objOU.MoveHere "LDAP://cn=user1,cn=users,dc=wisesoft,dc=co,dc=uk","cn=user1new"