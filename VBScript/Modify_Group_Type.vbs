Const ADS_GROUP_TYPE_GLOBAL_GROUP = &h2
Const ADS_GROUP_TYPE_LOCAL_GROUP = &h4
Const ADS_GROUP_TYPE_UNIVERSAL_GROUP = &h8
Const ADS_GROUP_TYPE_SECURITY_ENABLED = &h80000000
 
Set objGroup = GetObject _
    ("LDAP://cn=Students,cn=users,dc=wisesoft,dc=co,dc=uk") 
 
objGroup.Put "groupType", _
    ADS_GROUP_TYPE_UNIVERSAL_GROUP + ADS_GROUP_TYPE_SECURITY_ENABLED
objGroup.SetInfo
