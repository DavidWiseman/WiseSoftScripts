dim objTemplateUser
'<<<< Bind to a 'template' user object using the distinguished name >>>>
set objTemplateUser = GetObject("LDAP://cn=template,cn=users,dc=wisesoft,dc=co,dc=uk") 

dim objUser
'<<<< Bind to the user object using the distinguished name >>>>
Set objUser = GetObject("LDAP://cn=test.3,cn=users,dc=wisesoft,dc=co,dc=uk") 

'<<<< Copy logon hours from the template user >>>>
objUser.Put "logonHours", objTemplateUser.Get("logonHours")
objUser.SetInfo