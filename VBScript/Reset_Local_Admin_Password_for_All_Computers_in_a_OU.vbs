option explicit

DIM strSearchRoot,strSearchFilter,strSearchScope,strUser,strPassword
DIM cmd,cn,rs

' ******************* Setup ******************* 

' Specify the ou where the computer accounts are located
strSearchRoot = "cn=computers,dc=wisesoft,dc=org,dc=uk"
' Filter for all computers in the domain (modify if required)
strSearchFilter = "(objectCategory=Computer)"
' Child OUs are included in the search.  Change to "onelevel" to exclude child OUs.
strSearchScope = "subtree"
' Reset password for this local user account
strUser = "Administrator"
' Leave blank to auto-generate.  See generatePassword function.
strPassword = ""

' ********************************************** 

set cmd = createobject("ADODB.Command")
set cn = createobject("ADODB.Connection")
set rs = createobject("ADODB.Recordset")

cn.open "Provider=ADsDSOObject;"
cmd.activeconnection = cn

cmd.commandtext = "<LDAP://" & strSearchRoot & ">;" & strSearchFilter & ";" & _
		  "name;" & strSearchScope
'**** Bypass 1000 record limitation ****
cmd.properties("page size")=1000

set rs = cmd.execute

while rs.eof <> true and rs.bof <> true
	resetLocalUserPassword rs(0),strUser,strPassword
	rs.movenext
wend

cn.close

wscript.echo "Finished"

private function resetLocalUserPassword(BYVAL strComputer,BYVAL strUser,BYVAL strPassword)
	DIM objUser
	on error resume next

	if strPassword = "" then
		strPassword = generatePassword(strComputer)
	end if

	set objUser = getobject("WinNT://" & strComputer & "/" & strUser & ",user")
	
	' Check if we connected to the user object successfully
	if err.Number <> 0 then
		' Display an error message & clear the error
		wscript.echo "Error connecting to user '" &  strComputer & "\" & strUser & "'"
		err.Clear
	else
		' Change the password
		objUser.SetPassword strPassword
		objUser.SetInfo ' Save Changes
		if err.Number <> 0 then
			wscript.echo "Changed password for user '" &  strComputer & "\" & strUser & " to '" & strPassword & "'"
		else
			wscript.echo "Error Changing password for user '" &  strComputer & "\" & strUser
			err.Clear
		end if
	end if
	on error goto 0

end function

' Function to generate a different password for each computer.
' This is a simple function that has a static password component 
' followed by the name of the computer in lower case.
' This function should be changed to suit your own password policy
private function generatePassword(BYVAL strComputer)
	generatePassword = "AdminPa$$word@" & LCASE(strComputer)
end function