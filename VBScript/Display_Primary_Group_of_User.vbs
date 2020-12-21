'<<<< Prompt for a username >>>>
userName = inputbox("Please enter a username:")
if username = "" then wscript.quit

'<<<< Get the default naming context (saves hard-coding the domain name) >>>>
set objRoot = getobject("LDAP://RootDSE")
defaultNC = objRoot.get("defaultnamingcontext")

'<<<< Find the full path to the user object >>>>
userDN = FindUser(userName,defaultNC)
if userDN = "Not Found" then
	wscript.echo "User not found!"
	wscript.quit
end if

'<<<< Bind to the user object >>>>
set objUser = getobject(userDN)

'<<<< Get The primaryGroupID for the user >>>>
intPrimaryGroupID = objUser.Get("primaryGroupID")

'<<<< Find the group from the primaryGroupID >>>>
GroupName = getGroupFromToken(intPrimaryGroupID, defaultNC)

wscript.echo "Primary Group for " & username & " : " & GroupName


'<<<< Returns the group name from the primaryGroupToken.  The primaryGroupToken >>>>
'<<<< is calculated & therefore cannot be searched.  We must enumerate all groups >>>>
'<<<< to find the group we are searching for. The following search filter would not work: >>>>
'<<<< (&(objectCategory=Group)(primaryGroupToken=?)) >>>>
function getGroupFromToken(ByVal intPrimaryGroupID, searchRoot)

	Set cn = CreateObject("ADODB.Connection")
	Set cmd = CreateObject("ADODB.Command")

	cn.Open "Provider=ADsDSOObject;"
	cmd.ActiveConnection = cn
	cmd.CommandText = "<LDAP://" & searchRoot & ">;(objectCategory=Group);" & _
        "distinguishedName,primaryGroupToken,sAMAccountName;subtree"  
	Set rs = cmd.Execute
  
	Do Until rs.EOF
    		If rs("primaryGroupToken") = intPrimaryGroupID Then
        		getGroupFromToken = rs("sAMAccountName")
			exit do
    		End If
    		rs.MoveNext
	Loop
	cn.close
 
end function

'<<<< Returns an ADsPath for a user from the username >>>>
Function FindUser(Byval UserName, Byval Domain) 
	on error resume next

	set cn = createobject("ADODB.Connection")
	set cmd = createobject("ADODB.Command")
	set rs = createobject("ADODB.Recordset")

	cn.open "Provider=ADsDSOObject;"
	
	cmd.activeconnection=cn
	cmd.commandtext="SELECT ADsPath FROM 'LDAP://" & Domain & _
			"' WHERE sAMAccountName = '" & UserName & "'"
	
	set rs = cmd.execute

	if err<>0 then
		FindUser="Error connecting to Active Directory Database:" & err.description
		wscript.quit
	else
		if not rs.BOF and not rs.EOF then
     			rs.MoveFirst
     			FindUser = rs(0)
		else
			FindUser = "Not Found"
		end if
	end if
	cn.close
end function