' Get the NETBIOS Domain name
Set objSystemInfo = CreateObject("ADSystemInfo") 
strDomain = objSystemInfo.DomainShortName

' Prompt for userName
strUser = inputbox("Please enter the username (sAMAccountName):")

wscript.echo GetUserDN(strUser,strDomain)

Function GetUserDN(byval strUserName,byval strDomain)
	' Use name translate to return the distinguished name
	' of a user from the NT UserName (sAMAccountName)
	' and the NETBIOS domain name.

	Set objTrans = CreateObject("NameTranslate")
	objTrans.Init 1, strDomain
	objTrans.Set 3, strDomain & "\" & strUserName 
	strUserDN = objTrans.Get(1) 
	GetUserDN = strUserDN

end function