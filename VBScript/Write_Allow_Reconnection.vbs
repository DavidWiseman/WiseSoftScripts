Const TS_SESSION_ANY_CLIENT = 0
Const TS_SESSION_ORIGINATING = 1

Dim objUser
'<<<< Bind to the user object using the distinguished name >>>>
set objUser = getobject("LDAP://cn=test.3,cn=users,dc=wisesoft,dc=co,dc=uk")

' From any client"
objUser.ReconnectionAction = TS_SESSION_ANY_CLIENT '* From any client"
objUser.setinfo

' From originating client only
objUser.ReconnectionAction = TS_SESSION_ORIGINATING '* From originating client only
objUser.setinfo