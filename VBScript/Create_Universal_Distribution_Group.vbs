Const ADS_GROUP_TYPE_UNIVERSAL_GROUP = &h8

Set objOU = GetObject("LDAP://cn=users,dc=wisesoft,dc=co,dc=uk")
Set objGroup = objOU.Create("Group", "cn=MyUniversalDistributionGroup1")

objGroup.Put "sAMAccountName", "MyUniversalDistributionGroup1"
objGroup.Put "groupType", ADS_GROUP_TYPE_UNIVERSAL_GROUP
objGroup.SetInfo