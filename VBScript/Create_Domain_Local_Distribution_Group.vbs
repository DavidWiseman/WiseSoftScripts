Const ADS_GROUP_TYPE_LOCAL_GROUP = &h4

Set objOU = GetObject("LDAP://cn=users,dc=wisesoft,dc=co,dc=uk")
Set objGroup = objOU.Create("Group", "cn=MyLocalDistributionGroup1")

objGroup.Put "sAMAccountName", "MyLocalDistributionGroup1"
objGroup.Put "groupType", ADS_GROUP_TYPE_LOCAL_GROUP
objGroup.SetInfo