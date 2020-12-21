Const TS_SESSION_DISCONNECT = 0
Const TS_SESSION_END = 1
Const TS_SESSION_ANY_CLIENT = 0
Const TS_SESSION_ORIGINATING = 1

dim objUser
'<<<< Bind to the user object using the distinguished name >>>>
set objUser = getobject("LDAP://cn=test.3,cn=users,dc=wisesoft,dc=co,dc=uk")

'<<<< End a disconnected session >>>>
objUser.MaxDisconnectionTime = 0 '* Never
objUser.MaxDisconnectionTime = 30 '* 30 minutes

'<<<< Active session limit >>>>
objUser.MaxConnectionTime = 0 '* Never
objUser.MaxConnectionTime = 1440 '* 1 Day

'<<<< Idle session limit >>>>
objUser.MaxIdleTime = 0 'Never
objUser.MaxIdleTime = 2880 '* 2 Day


'<<<< When a session limit is reached or a connection is broken >>>>
objUser.BrokenConnectionAction = TS_SESSION_DISCONNECT '* Disconnect from session
objUser.BrokenConnectionAction = TS_SESSION_END '* End session

'<<<< Allow reconnection >>>>
objUser.ReconnectionAction = TS_SESSION_ANY_CLIENT '* From any client"
objUser.ReconnectionAction = TS_SESSION_ORIGINATING '* From originating client only

'<<<< Save Changes >>>>
objUser.setinfo