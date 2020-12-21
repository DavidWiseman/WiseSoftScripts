Const ADS_GROUP_TYPE_GLOBAL_GROUP = &h2

Set objOU = GetObject("LDAP://cn=users,dc=wisesoft,dc=co,dc=uk")
Set objGroup = objOU.Create("Group", "cn=MyGlobalDistributionGroup1")

objGroup.Put "sAMAccountName", "MyGlobalDistributionGroup1"
objGroup.Put "groupType", ADS_GROUP_TYPE_GLOBAL_GROUP
objGroup.SetInfo