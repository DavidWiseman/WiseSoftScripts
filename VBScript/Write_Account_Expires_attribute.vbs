'<<<< Bind to the user object using the distinguished name >>>>
set objUser = getobject("LDAP://cn=test.3,cn=users,dc=wisesoft,dc=co,dc=uk")

'<<<< Write Account Expiration Date >>>>>
objUser.AccountExpirationDate = "09/09/2005"
objUser.setinfo

'<<<< Don't expire account >>>>>
objUser.AccountExpirationDate = 0
objUser.setinfo