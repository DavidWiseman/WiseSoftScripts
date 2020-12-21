dim objUser

'<<<< Bind to the user object using the distinguished name >>>>
set objUser = getobject("LDAP://cn=test.3,cn=users,dc=wisesoft,dc=co,dc=uk")

'<<<< Enable User must change password at next logon >>>>
objUser.put "pwdLastSet", 0
objuser.setinfo

'<<<< Disable User must change password at next logon >>>>
objuser.put "pwdlastset", -1
objuser.setinfo