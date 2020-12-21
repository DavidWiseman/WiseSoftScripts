' Get the NETBIOS Domain name
Set objSystemInfo = CreateObject("ADSystemInfo") 
strDomain = objSystemInfo.DomainShortName

' Prompt for userName
strUserName = inputbox("Please enter the username (sAMAccountName) of the user to delete:")
if strUserName = "" then wscript.quit

' Call function to delete user
DeleteUser strUserName,strDomain

sub DeleteUser(byval strUserName,strDomain)
	' Function to delete a user account.
	' Use GetUserDN to convert username to distinguished name.
	' Use DN to bind to user object.  Get the container object
	' for the use (OU) and call the Delete method of the containter
	' object, passing the users common-name as a parameter.
	
	userDN = GetUserDN(strUserName,strDomain)
	set objUser = getObject("LDAP://" & userDN)
	set objContainer = getobject(objUser.Parent)
	objContainer.Delete "user","cn=" & objUser.cn

end sub

Function GetUserDN(byval strUserName,byval strDomain)
	' Use name translate to return the distinguished name
	' of a user from the NT UserName (sAMAccountName)
	' and the NETBIOS domain name.
	' e.g. cn=user1,cn=users,dc=wisesoft,dc=co,dc=uk

	Set objTrans = CreateObject("NameTranslate")
	objTrans.Init 1, strDomain
	objTrans.Set 3, strDomain & "\" & strUserName 
	strUserDN = objTrans.Get(1) 
	GetUserDN = strUserDN

end function