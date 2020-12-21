OPTION EXPLICIT
DIM objNetwork
DIM computerName
DIM ou

' Get the computerName of PC
set objNetwork = createobject("Wscript.Network")
computerName = objNetwork.ComputerName

' Call function to find OU from computer name
ou = getOUByComputerName(computerName)

wscript.echo ou


function getOUByComputerName(byval computerName)
	' *** Function to find ou/container of computer object from computer name ***
	
	DIM namingContext, ldapFilter, ou
	DIM cn, cmd, rs
	DIM objRootDSE
	
	' Bind to the RootDSE to get the default naming context for
	' the domain.  e.g. dc=wisesoft,dc=co,dc=uk
	set objRootDSE = getobject("LDAP://RootDSE")
	namingContext = objRootDSE.Get("defaultNamingContext")
	set objRootDSE = nothing

	' Construct an ldap filter to search for a computer object
	' anywhere in the domain with a name of the value specified.
	ldapFilter = "<LDAP://" & namingContext & _
 	">;(&(objectCategory=Computer)(name=" & computerName & "))" & _
	";distinguishedName;subtree"

	' Standard ADO code to query database
	set cn = createobject("ADODB.Connection")
	set cmd = createobject("ADODB.Command")

	cn.open "Provider=ADsDSOObject;"
	cmd.activeconnection = cn
	cmd.commandtext = ldapFilter
	
	set rs = cmd.execute

	if rs.eof <> true and rs.bof <> true then
		ou = rs(0)
		' Convert distinguished name into OU.
		' e.g. cn=CLIENT01,OU=WiseSoft_Computers,dc=wisesoft,dc=co,dc=uk
		' to: OU=WiseSoft_Computers,dc=wisesoft,dc=co,dc=uk
		ou = mid(ou,instr(ou,",")+1,len(ou)-instr(ou,","))
		getOUByComputerName = ou

	end if

	rs.close
	cn.close

end function