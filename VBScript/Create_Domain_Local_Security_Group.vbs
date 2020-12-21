Const ADS_GROUP_TYPE_LOCAL_GROUP = &h4
Const ADS_GROUP_TYPE_SECURITY_ENABLED = &h80000000

Set objOU = GetObject("LDAP://cn=users,dc=wisesoft,dc=co,dc=uk")
Set objGroup = objOU.Create("Group", "cn=MyLocalSecurityGroup1")

objGroup.Put "sAMAccountName", "MyLocalSecurityGroup1"
objGroup.Put "groupType", ADS_GROUP_TYPE_LOCAL_GROUP Or _
			ADS_GROUP_TYPE_SECURITY_ENABLED
objGroup.SetInfo