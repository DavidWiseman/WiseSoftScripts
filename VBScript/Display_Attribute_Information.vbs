' Prompt for attribute name. e.g. sn, givenName
strAttribute = inputbox("Please enter the name of the attribute:")
if strAttribute = "" then wscript.quit

' Bind to Root DSE (for schema naming context)
set objRoot = getobject("LDAP://RootDSE")

set cn = createobject("ADODB.Connection")
set cmd = createobject("ADODB.Command")
cn.open "Provider=ADsDSOObject;"
cmd.activeconnection = cn
cmd.commandtext = "<LDAP://" & objRoot.Get("schemaNamingContext") & _
	">;(&(objectclass=Attributeschema)(LDAPDisplayName=" & strAttribute & "));" & _
	"isSingleValued,rangeLower,rangeUpper,attributeSyntax;onelevel"

set rs = cmd.execute

while rs.EOF <> True and rs.BOF <> True
	' Get ADSTYPE and syntax name from attributeSyntax col
	select case rs(3)
	case "2.5.5.8"
		adsType = "ADSTYPE_BOOLEAN"
		syntaxName = "Boolean"	
	case "2.5.5.9"
		adsType = "ADSTYPE_INTEGER"
		syntaxName = "Integer/Enumeration"
	case "2.5.5.16"
		adsType = "ADSTYPE_LARGE_INTEGER"
		syntaxName = "Integer8"
	case "2.5.5.3"
		adsType = "ADSTYPE_CASE_EXACT_STRING"
		syntaxName = "CaseExactString"
	case "2.5.5.4"
		adsType = "ADSTYPE_CASE_IGNORE_STRING"
		syntaxName = "CaseIgnoreString"
	case "2.5.5.12"
		adsType = "ADSTYPE_CASE_IGNORE_STRING"
		syntaxName = "DirectoryString"
	case "2.5.5.5"
		adsType = "ADSTYPE_PRINTABLE_STRING"
		syntaxName = "IA5String"
	case "2.5.5.15"
		adsType = "ADSTYPE_NT_SECURITY_DESCRIPTOR"
		stntaxName = "NTSecurityDescriptor"
	case "2.5.5.6"
		adsType = "ADSTYPE_NUMERIC_STRING"
		syntaxName = "NumericString"
	case "2.5.5.10"
		adsType = "ADSTYPE_OCTET_STRING"
		syntaxName = "OctetString"
	case "2.5.5.2"
		adsType = "ADSTYPE_CASE_IGNORE_STRING"
		syntaxName = "OID"
	case "2.5.5.5"
		adsType = "ADSTYPE_PRINTABLE_STRING"
		syntaxName = "PrintableString"
	case "2.5.5.17"
		adsType = "ADSTYPE_OCTET_STRING"
		syntaxName = "SID"
	case "2.5.5.11"
		adsType = "ADSTYPE_UTC_TIME"
		syntax = "GeneralizedTime/UTCTime"
	case "2.5.5.14"
		adsType = ""
		syntaxName = "AccessPointDN"
	case "2.5.5.1"
		adsType = "ADSTYPE_DN_STRING"
		syntaxName = "DN"
	case "2.5.5.17"
		adsType = "ADSTYPE_DN_WITH_BINARY"
		syntaxName = "DNWithBinary"
	case "2.5.5.14"
		adsType = "ADSTYPE_DN_WITH_STRING"
		syntaxName = "DNWithString"
	case "2.5.5.7"
		adsType = ""
		syntaxName = "ORName"
	case "2.5.5.13"
		adsType = "ADSTYPE_CASE_IGNORE_STRING"
		syntaxName = "PresentationAddress"
	case "2.5.5.10"
		adsType = "ADSTYPE_OCTET_STRING"
		syntaxName = "ReplicaLink"
	case else
		adsType = "Unknown"
		syntaxName = "Unknown"
	end select
		
	' Display Attribute Details
	wscript.echo "LDAPDisplayName: " & strAttribute & vbcrlf & _
		"Single Valued : " & rs(0) & vbcrlf & _
		"Min Length: " & rs(1) & vbcrlf & _
		"Max Length: " & rs(2) & vbcrlf & _
		"Syntax: " & rs(3) & " (" & syntaxName & ")" & vbcrlf & _
		"ADSTYPE: " & adsType
	rs.movenext
wend