dim groupPath
dim userPath
groupPath = "LDAP://cn=ITC Staff,cn=users,dc=wisesoft,dc=co,dc=uk"
userPath = "LDAP://cn=test.3,cn=users,dc=wisesoft,dc=co,dc=uk"

removeFromGroup userPath,groupPath

sub removeFromGroup(userPath, groupPath)

	dim objGroup
	set objGroup = getobject(groupPath)
	
	for each member in objGroup.members
		if lcase(member.adspath) = lcase(userPath) then
			objGroup.Remove(userPath)
			exit sub
		end if
	next
	
end sub