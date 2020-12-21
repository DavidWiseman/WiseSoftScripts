OPTION EXPLICIT ' Variables must be declared
' *************************************************
' * Instructions
' *************************************************

' Edit the variables in the "Setup" section as required.
' Run this script from a command prompt in cscript mode.
' e.g. cscript usermod.vbs
' You can also choose to output the results to a text file:
' cscript usermod.csv >> results.txt

' *************************************************
' * Constants / Decleration
' *************************************************
Const adOpenStatic = 3
Const adLockOptimistic = 3
Const adCmdText = &H0001
Const ADS_PROPERTY_CLEAR = 1

DIM strSearchAttribute 
DIM strCSVHeader, strCSVFile, strCSVFolder
DIM strAttribute, userPath
DIM userChanges
DIM cn,cmd,rs
DIM objUser
DIM oldVal, newVal
DIM objField
DIM blnSearchAttributeExists
' *************************************************
' * Setup
' *************************************************

' The Active Directory attribute that is to be used to match rows in the CSV file to
' Active Directory user accounts.  It is recommended to use unique attributes.
' e.g. sAMAccountName (Pre Windows 2000 Login) or userPrincipalName
' Other attributes can be used but are not guaranteed to be unique.  If multiple user 
' accounts are found, an error is returned and no update is performed.
strSearchAttribute = "sAMAccountName" 'User Name (Pre Windows 2000)

' Folder where CSV file is located 
strCSVFolder = "C:\"
' Name of the CSV File
strCSVFile = "usermod.csv"

' *************************************************
' * End Setup
' *************************************************

' Setup ADO Connection to CSV file
Set cn = CreateObject("ADODB.Connection")
Set rs = CreateObject("ADODB.Recordset")

cn.Open "Provider=Microsoft.Jet.OLEDB.4.0;" & _
          "Data Source=" & strCSVFolder & ";" & _
          "Extended Properties=""text;HDR=YES;FMT=Delimited"""

rs.Open "SELECT * FROM [" & strCSVFile & "]", _
          cn, adOpenStatic, adLockOptimistic, adCmdText

' Check if search attribute exists
blnSearchAttributeExists=false
for each objField in rs.Fields
	if UCASE(objField.Name) = UCASE(strSearchAttribute) then
		blnSearchAttributeExists=true
	end if
Next
		
if blnSearchAttributeExists=false then
	MsgBox "'" & strSearchAttribute & "' attribute must be specified in the CSV header." & _
		VbCrLf & "The attribute is used to map the data the csv file to users in Active Directory.",vbCritical
	wscript.quit
end if

' Read CSV File
Do Until rs.EOF
	' Get the ADsPath of the user by searching for a user in Active Directory on the search attribute
	' specified, where the value is equal to the value in the csv file.
	' e.g. LDAP://cn=user1,cn=users,dc=wisesoft,dc=co,dc=uk
	userPath = getUser(strSearchAttribute,rs(strSearchAttribute))
	' Check that an ADsPath was returned
	if LEFT(userPath,6) = "Error:" then
		wscript.echo userPath
	else
		wscript.echo userPath
		' Get the user object
		set objUser = getobject(userpath)
		userChanges = 0
		' Update each attribute in the CSV string
		for each objField in rs.Fields
			strAttribute = objField.Name
			oldval = ""
			newval = ""
			' Ignore the search attribute (this is used only to search for the user account)
			if UCASE(strAttribute) <> UCASE(strSearchAttribute) and UCASE(strAttribute) <> "NULL" then
				newVal = rs(strAttribute) ' Get new attribute value from CSV file
				if ISNULL(newval) then
					newval = ""
				end If
				' Special handling for common-name attribute. If the new value contains
				' commas they must be escaped with a forward slash.
				If strAttribute = "cn" then
					newVal = REPLACE(newVal,",","\,")
				end If
				' Read the current value before changing it
				readAttribute strAttribute
								
				' Check if the new value is different from the update value
				if oldval <> newval then
					wscript.echo "Change " & strAttribute