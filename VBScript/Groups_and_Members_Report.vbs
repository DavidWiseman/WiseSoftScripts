option explicit ' Must Declare Variables
' Constants
Const ADS_GROUP_TYPE_GLOBAL_GROUP = &H2
Const ADS_GROUP_TYPE_DOMAIN_LOCAL_GROUP = &H4
Const ADS_GROUP_TYPE_UNIVERSAL_GROUP = &H8
Const ADS_GROUP_TYPE_SECURITY_ENABLED = &H80000000
' Declare Variables
dim cn,cmd, rs
dim ldapQuery,searchRoot,q,fileName
dim objRoot,objFSO,objFile
q = """"

' Report FileName:
fileName = "WiseSoft_Groups_And_Group_Members_Report.csv"

' Create the report csv/text file
set objFSO = createobject("scripting.filesystemobject")
set objFile = objFSO.createtextfile(fileName)

' Get the default naming context. e.g. DC=wisesoft,DC=co,DC=UK
set objRoot = getobject("LDAP://rootDSE")
searchRoot = objRoot.Get("defaultNamingContext")

set cn = createobject("ADODB.Connection")
set cmd = createobject("ADODB.Command")

cn.open "Provider=ADsDSOObject;"
cmd.activeconnection = cn
' Query to return all the groups in the current domain
cmd.commandtext = "<LDAP://" & searchRoot & ">;(objectCategory=Group);ADsPath,primaryGroupToken;subtree"
cmd.Properties("Page Size") = 1000

set rs = cmd.execute

' Write the CSV header line
objFile.WriteLine("""Group Type"",""Group"",""Group Path"",""Member Type"",""Member"",""Member Path""")

' For each group...
while rs.eof <> true and rs.bof <> true
	dim objGroup, objMember, memberName, groupType
	memberName = ""
	' Bind to the group object
	set objGroup = GetObject(rs(0))
	' Get the Group Type
	SELECT CASE objGroup.GroupType
	CASE ADS_GROUP_TYPE_GLOBAL_GROUP OR ADS_GROUP_TYPE_SECURITY_ENABLED
		groupType = "Global Security"
	CASE ADS_GROUP_TYPE_DOMAIN_LOCAL_GROUP OR ADS_GROUP_TYPE_SECURITY_ENABLED
		groupType = "Domain Local Security"
	CASE ADS_GROUP_TYPE_UNIVERSAL_GROUP OR ADS_GROUP_TYPE_SECURITY_ENABLED
		groupType =  "Universal Security"
	CASE ADS_GROUP_TYPE_GLOBAL_GROUP
		groupType = "Global Distribution"
	CASE ADS_GROUP_TYPE_DOMAIN_LOCAL_GROUP
		groupType = "Domain Local Distribution"
	CASE ADS_GROUP_TYPE_UNIVERSAL_GROUP
		groupType = "Universal Distribution"
	CASE ELSE
		groupType = "BuiltIn"
	END SELECT
	' List primary group members
	reportPimaryGroupMembers groupType,objGroup.ADsPath,objGroup.sAMAccountName,rs(1)

	' List Group members
	for each objMember in objGroup.Members
		If objMember.Class = "user" then
			memberName = objMember.sAMAccountName
		elseif objMember.Class ="group" then
			memberName = objMember.sAMAccountName
		else
			memberName = objMember.Name
		end if
		' Write line to csv report
		objFile.WriteLine(q & groupType & q & "," & q & objGroup.sAMAccountName & q & "," & q & objGroup.ADsPath & q & "," & q & _
				objMember.Class & q & "," & q & memberName & q & "," & q & objmember.ADsPath & q)
	next
	
	rs.movenext
wend

rs.close
cn.close
objFile.Close

sub reportPimaryGroupMembers(byval groupType,byval groupPath,byval groupName,byval primaryGroupToken)
	' Procedure to list primary group members
	' Search for users where the primarygroupid equals the PrimayGroupToken passed to the procedure
	dim cnPrimary,cmdPrimary,rsPrimary
	set cnPrimary = createobject("ADODB.Connection")
	set cmdPrimary = createobject("ADODB.Command")
	set rsPrimary = createobject("ADODB.Recordset")
	
	cnPrimary.open "Provider=ADsDSOObject;"
	cmdPrimary.activeconnection=cn

	
	cmdPrimary.commandtext = "<LDAP://" & searchRoot & ">;(&(objectCategory=Person)(objectClass=User)(PrimaryGroupID=" & primaryGroupToken & "));ADsPath,sAMAccountName;subtree"
	cmdPrimary.properties("Page Size") = 1000
	set rsPrimary = cmdPrimary.execute

	while rsPrimary.eof<>true and rsPrimary.bof<>true
		' Write line to csv report
		objFile.WriteLine(q & groupType & q & "," & q & groupName & q & "," & q & groupPath & q & "," & q & _
				"user" & q & "," & q & rsPrimary(1) & q & "," & q & rsPrimary(0) & q)

		rsPrimary.movenext
	wend
	rsPrimary.Close
	cnPrimary.close
	
end sub