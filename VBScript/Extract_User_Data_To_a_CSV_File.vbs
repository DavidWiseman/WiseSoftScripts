OPTION EXPLICIT

dim FileName, multivaluedsep,strAttributes
dim strFilter, strRoot, strScope
dim cmd, rs,cn
dim objRoot, objFSO,objCSV
dim comma, q, i, j, mvsep, strAttribute, strValue

' ********************* Setup *********************

' The filename of the csv file produced by this script
FileName ="userexport.csv"
' Seperator used for multi-valued attributes
multivaluedsep = ";"
' comma seperated list of attributes to export
strAttributes = "sAMAccountName,givenName,initials,sn,displayName,description,physicalDeliveryOfficeName," & _
		"telephoneNumber,mail,wWWHomePage,cn"

' Default filter for all user accounts (ammend if required)
strFilter = "(&(objectCategory=person)(objectClass=user))"
' scope of search (default is subtree - search all child OUs)
strScope = "subtree"
' search root. e.g. ou=MyUsers,dc=wisesoft,dc=co,dc=uk
' leave blank to search from domain root
strRoot = ""

' *************************************************

q = """"

set cmd = createobject("ADODB.Command")
set cn = createobject("ADODB.Connection")
set rs = createobject("ADODB.Recordset")

cn.open "Provider=ADsDSOObject;"
cmd.activeconnection = cn

if strRoot = "" then
	set objRoot = getobject("LDAP://RootDSE")
	strRoot = objRoot.get("defaultNamingContext") 
end if

cmd.commandtext = "<LDAP://" & strRoot & ">;" & strFilter & ";" & _
		  strAttributes & ";" & strScope

'**** Bypass 1000 record limitation ****
cmd.properties("page size")=1000

set rs = cmd.execute
set objFSO = createobject("Scripting.FileSystemObject")
set objCSV = objFSO.createtextfile(FileName)

comma = "" ' first column does not require a preceding comma
i = 0 

' create a header row and count the number of attributes
for each strAttribute in SPLIT(strAttributes,",")
	objcsv.write(comma & q & strAttribute & q)
	comma = "," ' all columns apart from the first column require a preceding comma
	i = i + 1
next

' for each item returned by the Active Directory query
while rs.eof <> true and rs.bof <> true
	comma="" ' first column does not require a preceding comma
	objcsv.writeline ' Start a new line
	' For each column in the result set
	for j = 0 to (i - 1)
		select case typename(rs(j).value)
		case "Null" ' handle null value
			objcsv.write(comma & q & q)
		case "Variant()" ' multi-valued attribute
			' Multi-valued attributes will be seperated by value specified in
			' "multivaluedsep" variable
			mvsep = "" 'No seperator required for first value
			objcsv.write(comma & q)
			for each strValue in rs(j).Value
				' Write value
				' single double quotes " are replaced by double double quotes ""
				objcsv.write(mvsep & replace(strValue,q,q & q))
				mvsep = multivaluedsep ' seperator used when more than one value returned
			next
			objcsv.write(q)
		case else
			' Write value
			' single double quotes " are replaced by double double quotes ""
			objcsv.write(comma & q & replace(rs(j).value,q,q & q) & q)
		end select
		
		comma = "," ' all columns apart from the first column require a preceding comma
	next
	rs.movenext
wend

' Close csv file and ADO connection
cn.close
objCSV.Close

wscript.echo "Finished"