OPTION Explicit

Dim strGroupDN,strNETBIOSDomain, strGroupName
Dim objSystemInfo, objGroup, objMember
Dim strAttributeName, strAttributeValue, strValue
Dim strMemberDN
' ********************* Setup *********************

' Group to remove users from
strGroupName="test"
' Remove all users from group where the wWWHomePage value does not equal "20"
strAttributeName = "wWWHomePage"
strAttributeValue = "20"

' *************************************************

SET objSystemInfo = CREATEOBJECT("ADSystemInfo") 
'Required for name translate
strNETBIOSDomain = objSystemInfo.DomainShortName

' Convert group name to distinguished name
strGroupDN =  GetDN(strNETBIOSDomain,strGroupName)

SET objGroup = GETOBJECT("LDAP://" & strGroupDN)

FOR EACH objMember in objGroup.Members
		On Error Resume Next ' Ignore error that occurs when reading blank attribute value
		strValue = ""
		strValue = objMember.Get(strAttributeName)
		On Error GoTo 0
		' Remove user from group if attribute value does not match expected value
  		If strValue  <> strAttributeValue Then
  			WScript.Echo "Removing user " & objMember.ADsPath & " from group"
  			objGroup.Remove objMember.ADsPath
  		End If
Next

' Function to convert name into distinguished name format
FUNCTION GetDN(BYVAL strDomain,strObject)
	' Use name translate to return the distinguished name
	' of a user from the NT UserName (sAMAccountName)
	' and the NETBIOS domain name.
	DIM objTrans

	SET objTrans = CREATEOBJECT("NameTranslate")
	objTrans.Init 1, strDomain
	objTrans.SET 3, strDomain & "\" & strObject
	GetDN = objTrans.GET(1) 

END Function