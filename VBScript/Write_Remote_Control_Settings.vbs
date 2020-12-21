dim objUser
'<<<< Bind to the user object using the distinguished name >>>>
set objUser = getobject("LDAP://cn=test.3,cn=users,dc=wisesoft,dc=co,dc=uk")

' Remote Control - Disabled
objuser.EnableRemoteControl = 0
objUser.setinfo	

' Remote Control - Enabled
' Require User's permission - Enabled
' Level of Control - Interact with the session
objuser.EnableRemoteControl = 1
objUser.setinfo	

' Remote Control - Enabled
' Require Users permission - Disabled
' Level of Control - Interact with the session
objuser.EnableRemoteControl = 2
objUser.setinfo	

' Remote Control - Enabled
' Require Users permission - Enabled
' Level of Control - View the user's session
objuser.EnableRemoteControl = 3
objUser.setinfo	

' Remote Control - Enabled
' Require Users permission - Disabled
' Level of Control - View the user's session
objuser.EnableRemoteControl = 4
objUser.setinfo	