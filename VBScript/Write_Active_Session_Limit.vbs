Dim objUser
'<<<< Bind to the user object using the distinguished name >>>>
set objUser = getobject("LDAP://cn=test.3,cn=users,dc=wisesoft,dc=co,dc=uk")

' Never
objUser.MaxConnectionTime = 0 '* Never
objUser.setinfo

' 1 Day
objUser.MaxConnectionTime = 1440 '* 1 Day
objUser.setinfo