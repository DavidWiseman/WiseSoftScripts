' Dictionary object created here to be accessed by all functions
set objGroupList = CreateObject("Scripting.Dictionary")
' Ensure Exists method of dictionary object is not case sensitive
objGroupList.CompareMode = vbTextCompare

' Bind to a couple of user objects to test the group membership function
set objUser1 = getobject("LDAP://cn=user1,ou=Staff,dc=wisesoft,dc=co,dc=uk")
set objUser2 = getobject("LDAP://cn=user2,ou=Staff,dc=wisesoft,dc=co,dc=uk")

strGroup = "domain users"

if isMember(objUser1,strGroup) then
	wscript.echo objUser1.sAMAccountName & " is a member of " & strGroup
else
	wscript.echo objUser1.sAMAccountName & " is NOT a member of " & strGroup
end if
if isMember(objUser2,strGroup) then
	wscript.echo objUser2.sAMAccountName & " is a member of " & strGroup
else
	wscript.echo objUser2.sAMAccountName & " is NOT a member of " & strGroup
end if

' Groups have already been enumerated for user1 and user2 so subsequent calls to
' the isMember functions will use the cache in the dictionary object.  

strGroup = "Developers"

if isMember(objUser1,strGroup) then
	wscript.echo objUser1.sAMAccountName & " is a member of " & strGroup
else
	wscript.echo objUser1.sAMAccountName & " is NOT a member of " & strGroup
end if
if isMember(objUser2,strGroup) then
	wscript.echo objUser2.sAMAccountName & " is a member of " & strGroup
else
	wscript.echo objUser2.sAMAccountName & " is NOT a member of " & strGroup
end if

function isMember(byref objADObject,groupName)
	' Function to test the group membership of a user/computer
	' objADObject is user or computer object to test group membership
	' groupName is the name (sAMAccountName) of the group

	' Check if groups have been cached for this user/computer in the dictionary
	if not objGroupList.Exists(objADObject.sAMAccountName) then

		' Add user to dictionary
		objGroupList.add objADObject.sAMAccountName,""
		' Add primary group to dictionary (see GetPrimaryGroup function)
		objGroupList.add objADObject.sAMAccountName & "|" & GetPrimaryGroup(objADObject),""

		' For each group the user/computer is a member of
		For Each objGroup in objADObject.Groups
			if not objGroupList.Exists(objADObject.sAMAccountName & "|" & objGroup.sAMAccountName) then
				' Add group to the dictionary
				objGroupList.add objADObject.sAMAccountName & "|" & objGroup.sAMAccountName,""
				' Get nested groups (groups this group is a member of etc)
    				getNested objGroup,objADObject.sAMAccountName
			end if
		Next
		
	end if
	' All groups the user is a member of have now been cached in a dictionary object
	' (objGroupList).  Simply call the exists method on the dictionary object to check
	' the user/computer group membership.  Any subsequent calls to this function for the same
	' user/computer will use the dictionary object rather than re-enumerating the groups.
	
	isMember =  objGroupList.exists(objADObject.sAMAccountName & "|" & groupName)

end function 

Function GetNested(byref objGroup,byval sAMAccountName)
	' Find the groups that the input group (objGroup) is a member of
	' add these groups to the dictionary and call GetNested function 
	' again on these groups. (Recursive)

    	On Error Resume Next
    	colMembers = objGroup.GetEx("memberOf")
	on error goto 0

	if not isEmpty(colMembers) then
    		For Each strMember in colMembers
       			Set objNestedGroup = GetObject("LDAP://" & strMember)
			item = sAMAccountName & "|" & objNestedGroup.sAMAccountName
			if objGroupList.Exists(Item) = False then
				' Add group to the dictionary
				objGroupList.add item,""
				' Get groups that this group is a member of
				GetNested objNestedGroup,sAMAccountName
			end if
    		Next
	end if
End Function

Function GetPrimaryGroup(objADObject)
	' Function to find the primary group of a user/computer object
	' Search Active Directory using ADO to get a list of primary 
	' group tokens for each group.  Search results to find