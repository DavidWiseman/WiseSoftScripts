Const ADS_GROUP_TYPE_UNIVERSAL_GROUP = &h8
Const ADS_GROUP_TYPE_SECURITY_ENABLED = &h80000000

Set objOU = GetObject("LDAP://cn=users,dc=wisesoft,dc=co,dc=uk")
Set objGroup = objOU.Create("Group", "cn=MyUniversalSecurityGroup1")

objGroup.Put "sAMAccountName", "MyUniversalSecurityGroup1"
objGroup.Put "groupType", ADS_GROUP_TYPE_UNIVERSAL_GROUP Or _
    			ADS_GROUP_TYPE_SECURITY_ENABLED
objGroup.SetInfo