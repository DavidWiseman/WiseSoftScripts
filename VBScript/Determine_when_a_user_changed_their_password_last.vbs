Set objUser = GetObject("LDAP://cn=user1,cn=users,dc=wisesoft,dc=co,dc=uk")
Wscript.Echo "Password last changed: " & objUser.PasswordLastChanged