Const ADS_PROPERTY_CLEAR = 1

DIM objUser
'<<<< Bind to the user object using the distinguished name >>>>
set objUser = GetObject("LDAP://cn=user1,cn=users,dc=wisesoft,dc=org,dc=uk")

' Control access through remote access policy
objUser.PutEx ADS_PROPERTY_CLEAR, "msNPAllowDialIn", null
objUser.setinfo 'Save Changes
wscript.echo "Setting changed to 'Control access through remote access policy'"

' Allow Access
objUser.Put "msNPAllowDialIn", TRUE
objUser.setinfo 'Save Changes
wscript.echo "Setting changed to 'Allow Access'"

' Deny Access
objUser.Put "msNPAllowDialIn", FALSE
objUser.setinfo 'Save Changes
wscript.echo "Setting changed to 'Deny Access'"
