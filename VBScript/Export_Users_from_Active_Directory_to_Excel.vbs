' Bind to RootDSE - this object is used to 
' get the default configuration naming context
' e.g. dc=wisesoft,dc=co,dc=uk
set objRootDSE = getobject("LDAP://RootDSE")

' File name to export to
strExportFile = "C:\MyExport.xls" 
' Root of search set to default naming context.
' e.g. dc=wisesoft,dc=co,dc=uk
' RootDSE saves hard-coding the domain.  
' If want to search within an OU rather than the domain,
' specify the distinguished name of the ou.  e.g. 
' ou=students,dc=wisesoft,dc=co,dc=uk"
strRoot = objRootDSE.Get("DefaultNamingContext")
' Filter for user accounts - could be modified to search for specific users,
' such as those with mailboxes, users in a certain department etc.
strfilter = "(&(objectCategory=Person)(objectClass=User))"
' Attributes to return from the query
strAttributes = "sAMAccountName,userPrincipalName,givenName,sn," & _
		"initials,displayName,physicalDeliveryOfficeName," & _
		"telephoneNumber,mail,wWWHomePage,profilePath," & _
		"scriptPath,homeDirectory,homeDrive,title,department," & _
		"company,manager,homePhone,pager,mobile," & _
		"facsimileTelephoneNumber,ipphone,info," & _
		"streetAddress,postOfficeBox,l,st,postalCode,c"
'Scope of the search.  Change to "onelevel" if you didn't want to search child OU's
strScope = "subtree"

set cn = createobject("ADODB.Connection")
set cmd = createobject("ADODB.Command")

cn.open "Provider=ADsDSOObject;"
cmd.ActiveConnection = cn
cmd.commandtext = "<LDAP://" & strRoot & ">;" & strFilter & ";" & _
		   strAttributes & ";" & strScope

set rs = cmd.execute

' Use Excel COM automation to open Excel and create an excel workbook
set objExcel = CreateObject("Excel.Application")
set objWB = objExcel.Workbooks.Add
set objSheet = objWB.Worksheets(1)

' Copy Field names to header row of worksheet
For i = 0 To rs.Fields.Count - 1
	objSheet.Cells(1, i + 1).Value = rs.Fields(i).Name
	objSheet.Cells(1, i + 1).Font.Bold = True
Next

' Copy data to the spreadsheet
objSheet.Range("A2").CopyFromRecordset(rs)
' Save the workbook
objWB.SaveAs(strExportFile)

' Clean up
rs.close
cn.close
set objSheet = Nothing
set objWB =  Nothing
objExcel.Quit()
set objExcel = Nothing