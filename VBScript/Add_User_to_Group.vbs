dim groupPath
dim userPath

groupPath = "LDAP://cn=ITC Staff,cn=users,dc=wisesoft,dc=co,dc=uk"
userPath = "LDAP://cn=test.3,cn=users,dc=wisesoft,dc=co,dc=uk"

addToGroup userPath,groupPath

sub addToGroup(userPath, groupPath)
	dim objGroup
	set objGroup = getobject(groupPath)
	
	for each member in objGroup.members
		if lcase(member.adspath) = lcase(userPath) then
			exit sub
		end if
	next
	objGroup.Add(userPath)

end sub