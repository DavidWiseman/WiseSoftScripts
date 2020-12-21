Option Explicit
' Created By David Wiseman
' http://www.wisesoft.co.uk
' *************************************************
' * Description
' *************************************************
' Script to update group membership based on a CSV file
' *************************************************
' * Instructions
' *************************************************
' Requires a CSV file with the headers "UserName" and "Group"
' Edit the variables in the "Setup" section as required.
' Run this script from a command prompt in cscript mode.
' e.g. cscript groupmemberupdate.vbs
' You can also choose to output the results to a text file:
' cscript groupmemberupdate.vbs >> results.txt

' *************************************************
' *************************************************
' * Constants / Decleration
' *************************************************
CONST adOpenStatic = 3
CONST adLockOptimistic = 3
CONST adCmdText = &H0001
Const ADS_NAME_INITTYPE_DOMAIN = 1
Const ADS_NAME_TYPE_NT4 = 3
Const ADS_NAME_TYPE_1779 = 1

Dim strFolder,strCSVFile, strCSVFolder, strDomain
Dim strUserDN, strGroupDN, strUser,strGroup, strPreviousUser,strPreviousGroup
Dim cn,rs, objSystemInfo, objGroup
' *************************************************
' * Setup
' *************************************************
' Folder where CSV file is located 
strCSVFolder = "C:\"
' Name of the CSV File
strCSVFile = "groupmemberupdate.csv"

' *************************************************
' * End Setup
' *************************************************
' Get domain name
SET objSystemInfo = CREATEOBJECT("ADSystemInfo") 
strDomain = objSystemInfo.DomainShortName
' Setup ADO Connection to CSV file
SET cn = CREATEOBJECT("ADODB.Connection")
SET rs = CREATEOBJECT("ADODB.Recordset")

cn.Open "Provider=Microsoft.Jet.OLEDB.4.0;" & _
          "Data Source=" & strCSVFolder & ";" & _
          "Extended Properties=""text;HDR=YES;FMT=Delimited"""


rs.Open "SELECT * FROM [" & strCSVFile & "]", _
          cn, adOpenStatic, adLockOptimistic, adCmdText
             
' Read CSV File
DO Until rs.EOF
	' Read UserName/GroupName columns from CSV file
	strUser= rs("UserName")
	strGroup= rs("Group")
	' Translate username to distinguished name format. e.g. CN=david.wiseman,OU=MyUsers,DC=wisesoft,DC=co,DC=uk
	If strUser <> strPreviousUser Then ' Don't bother with name translate if user is same as previous row
		strUserDN =  GetDN(strUser,strDomain)
	End If
	' Translate group to distinguished name format. e.g. CN=IT Operations,OU=MyGroups,DC=wisesoft,DC=co,DC=uk
	If strGroup <> strPreviousGroup Then ' Don't bother with name translate if group is same as previous row
		strGroupDN = GetDN(strGroup,strDomain)
	End If
	' Bind to group object
	Set objGroup = GetObject("LDAP://" & strGroupDN)
	On Error Resume Next ' Ignore errors
	' Add user as member of group
	objGroup.Add "LDAP://" & strUserDN
	' Check if an error occurred (e.g. user is already a member of the group)
	select Case Err.Number
	Case -2147019886 
		WScript.Echo "'" & strUser & "' is already a member of '" & strGroup & "'"
	Case 0
		WScript.Echo "Added '" & strUser & "' to group '" & strGroup & "'"
	Case Else
		WScript.Echo "Error adding user '" & strUser & " to group '" & strGroup & "': " & Err.Number & " " & Err.Description
	End Select 
	Err.Clear
	On Error GoTo 0 ' Turn off resume next error handling
	' Store name of group/user we've just processed
	strPreviousUser = strUser
	strPreviousGroup = strGroup
	' Move to next row in CSV
	rs.MoveNext
Loop

' Get distinguished name from NT name. e.g. David.Wiseman >>> CN=david.wiseman,OU=MyUsers,DC=wisesoft,DC=co,DC=uk
' See http://www.rlmueller.net/NameTranslateFAQ.htm for more info on NameTranslate
FUNCTION GetDN(ByVal strName,ByVal strDomain)

	Dim objTrans, strDN
	SET objTrans = CREATEOBJECT("NameTranslate")
	objTrans.Init ADS_NAME_INITTYPE_DOMAIN, strDomain
	objTrans.SET ADS_NAME_TYPE_NT4, strDomain & "\" & strName 
	strDN = objTrans.GET(ADS_NAME_TYPE_1779) 
	GetDN = 