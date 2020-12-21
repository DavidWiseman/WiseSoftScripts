Option Explicit

dim FileName, multivaluedsep,strAttributes
dim strFilter, strRoot, strScope
dim cmd, rs,cn
Dim objRoot, objFSO,objCSV
Dim comma, q, i, mvsep, strAttribute, strValue,strSearchAttributes
dim objUser

' ********************* Setup *********************

' The filename of the csv file produced by this script
FileName ="userexport.csv"
' Seperator used for multi-valued attributes
multivaluedsep = ";"
' comma seperated list of attributes to export
strAttributes = "sAMAccountName,givenName,initials,sn,displayName,description,physicalDeliveryOfficeName," & _
		"telephoneNumber,mail,wWWHomePage,cn,terminalservicesprofilepath,terminalserviceshomedrive,terminalserviceshomedirectory,allowlogon"

' Default filter for all user accounts (ammend if required)
strFilter = "(&(objectCategory=person)(objectClass=user))"
' scope of search (default is subtree - search all child OUs)
strScope = "subtree"
' search root. e.g. ou=MyUsers,dc=wisesoft,dc=co,dc=uk
' leave blank to search from domain root
strRoot = ""

' *************************************************

q = """"
comma = "" ' first column does not require a preceding comma
i = 0 
Set objFSO = createobject("Scripting.FileSystemObject")
Set objCSV = objFSO.createtextfile(FileName)

' Create CSV header row and get attributes to use in search
For Each strAttribute In SPLIT(strAttributes,",")
	Select Case LCASE(strAttribute)
		Case "terminalservicesprofilepath","terminalserviceshomedrive","terminalserviceshomedirectory","allowlogon","manager_samaccountname"
			' Terminal services attributes are stored in the userparameters attribute and can be read individually
			' via the IADsTSUserEx interface. This requires us to bind to each user account returned by the search (slow)
			' Add the "adspath" attribute to allow us to bind to the user account where terminal services attributes are 
			' specified
			If INSTR(1,strSearchAttributes,"adspath",1) = 0 Then ' Check if we don't already have adspath attribute
				IF strSearchAttributes <> "" Then
					strSearchAttributes = strSearchAttributes & ","
				End If
				strSearchAttributes = strSearchAttributes & "adspath"
			End If
		Case Else
			' Append attribute to the search attributes
			If strSearchAttributes <> "" Then
				strSearchAttributes = strSearchAttributes & ","
			End If
			strSearchAttributes = strSearchAttributes & strAttribute
	END Select
	' Write CSV File Header
	objcsv.write(comma & q & strAttribute & q)
	comma = "," ' all columns apart from the first column require a preceding comma
	i = i + 1
Next

set cmd = createobject("ADODB.Command")
set cn = createobject("ADODB.Connection")
set rs = createobject("ADODB.Recordset")

cn.open "Provider=ADsDSOObject;"
cmd.activeconnection = cn

' If root = "", use the default naming context (current domain)
if strRoot = "" then
	set objRoot = getobject("LDAP://RootDSE")
	strRoot = objRoot.get("defaultNamingContext") 
end if

cmd.commandtext = "<LDAP://" & strRoot & ">;" & strFilter & ";" & _
		  strSearchAttributes & ";" & strScope

'**** Bypass 1000 record limitation ****
cmd.properties("page size")=1000

set rs = cmd.execute

' for each item returned by the Active Directory query
while rs.eof <> true and rs.bof <> True
	Set objUser = Nothing ' Used only for terminal services attributes
	comma="" ' first column does not require a preceding comma
	objcsv.writeline ' Start a new line
	' For each column in the result set
	for each strAttribute in SPLIT(strAttributes,",")
		select case strAttribute
			case "terminalservicesprofilepath"
				' Bind to user account if required (only bind once per user if more than 1 
				' terminal services attribute is specified)
				If objUser Is Nothing Then
					Set objUser = GETOBJECT(rs("adspath"))
				End If
				objCSV.Write(comma & q & replace(objUser.TerminalServicesProfilePath,q,q & q) & q)
			case "terminalserviceshomedrive"
				' Bind to user acco