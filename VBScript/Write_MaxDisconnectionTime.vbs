Dim objUser
'<<<< Bind to the user object using the distinguished name >>>>
set objUser = getobject("LDAP://cn=test.3,cn=users,dc=wisesoft,dc=co,dc=uk")

' Never
objUser.MaxDisconnectionTime = 0 '* Never
objUser.setinfo

' 30 Minutes
objUser.MaxDisconnectionTime = 30 '* 30 minutes
objUser.setinfo
