Dim objUser
'<<<< Bind to the user object using the distinguished name >>>>
set objUser = getobject("LDAP://cn=test.3,cn=users,dc=wisesoft,dc=co,dc=uk")

' Never
objUser.MaxIdleTime = 0 'Never
objUser.setinfo

' 2 Days
objUser.MaxIdleTime = 2880 '* 2 Days
objUser.setinfo