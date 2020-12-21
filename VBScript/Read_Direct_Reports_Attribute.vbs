set objUser = GetObject("LDAP://cn=user1,cn=users,dc=wisesoft,dc=co,dc=uk")

For Each strValue in objUser.directReports
    WScript.Echo "Direct Reports: " & strValue
Next