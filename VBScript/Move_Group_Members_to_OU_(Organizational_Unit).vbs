' *********************** Setup ***********************
strGroupPath = "LDAP://CN=MyGroup,OU=MyOU,DC=wisesoft,DC=co,DC=uk"
strNewOUPath = "LDAP://OU=NewOU,DC=wisesoft,DC=co,DC=uk"
*******************************************************

set objOU = GETOBJECT(strNewOUPath)
set objGroup = GETOBJECT(strGroupPath)

for each objMember in objGroup.Members
	objOU.MoveHere objMember.ADsPath,vbNullString
next