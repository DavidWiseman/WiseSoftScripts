' Constants required for name translate
Const ADS_NAME_INITTYPE_GC = 3
Const ADS_NAME_TYPE_NT4 = 3
Const ADS_NAME_TYPE_1779 = 1

'Get the NETBIOS name of the domain
Set objSystemInfo = CreateObject("ADSystemInfo") 
strDomain = objSystemInfo.DomainShortName

' Get the name of the computer
set objNetwork = createobject("Wscript.Network")
strComputer = objNetwork.ComputerName

' Call function to return the distinguished name (DN) of the computer
strComputerDN = getComputerDN(strComputer,strDomain)

wscript.echo strComputerDN


function getComputerDN(byval strComputer,byval strDomain)
	' Function to get the distinguished name of a computer
	' from the NETBIOS name of the computer (strcomputer)
	' and the NETBIOS name of the domain (strDomain) using
	' name translate

	Set objTrans = CreateObject("NameTranslate")
	' Initialize name translate using global catalog
	objTrans.Init ADS_NAME_INITTYPE_GC, ""
	' Input computer name (NT Format)
	objTrans.Set ADS_NAME_TYPE_NT4, strDomain & "\" & strComputer & "$"
	' Get Distinguished Name.
	getComputerDN = objTrans.Get(ADS_NAME_TYPE_1779)

end function