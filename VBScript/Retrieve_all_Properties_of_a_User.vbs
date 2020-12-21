' Declare Constants
Const ADSTYPE_DN_STRING = 1 
Const ADSTYPE_CASE_EXACT_STRING = 2 
Const ADSTYPE_CASE_IGNORE_STRING = 3 
Const ADSTYPE_PRINTABLE_STRING = 4 
Const ADSTYPE_NUMERIC_STRING = 5 
Const ADSTYPE_BOOLEAN = 6 
Const ADSTYPE_INTEGER = 7 
Const ADSTYPE_OCTET_STRING = 8 
Const ADSTYPE_UTC_TIME = 9 
Const ADSTYPE_LARGE_INTEGER = 10 
Const ADSTYPE_PROV_SPECIFIC = 11 
Const ADSTYPE_OBJECT_CLASS = 12 
Const ADSTYPE_CASEIGNORE_LIST = 13 
Const ADSTYPE_OCTET_LIST = 14 
Const ADSTYPE_PATH = 15 
Const ADSTYPE_POSTALADDRESS = 16 
Const ADSTYPE_TIMESTAMP = 17 
Const ADSTYPE_BACKLINK = 18 
Const ADSTYPE_TYPEDNAME = 19 
Const ADSTYPE_HOLD = 20 
Const ADSTYPE_NETADDRESS = 21 
Const ADSTYPE_REPLICAPOINTER = 22 
Const ADSTYPE_FAXNUMBER = 23 
Const ADSTYPE_EMAIL = 24 
Const ADSTYPE_NT_SECURITY_DESCRIPTOR = 25 
Const ADSTYPE_UNKNOWN = 26 

' Get the user object fetch the properties into the cache
set objUser = getobject("LDAP://cn=user1,cn=users,dc=wisesoft,dc=co,dc=uk")
objUser.GetInfo

' Enumerate the properties of the user object
for i = 0 to objUser.PropertyCount - 1
	' Get the name of the property
	strOutput = strOutput & objUser.Item(i).Name & vbcrlf
	
	' Enumerate each value of the property
	for each v in objUser.Item(i).Values
		select case objUser.Item(i).ADsType
			case ADSTYPE_DN_STRING
				strOutput = strOutput & vbtab & v.DNString & vbcrlf
			case ADSTYPE_CASE_EXACT_STRING
				strOutput = strOutput & vbtab & v.CaseExactString &vbcrlf
			case ADSTYPE_CASE_IGNORE_STRING
				strOutput = strOutput & vbtab & v.CaseIgnoreString & vbcrlf
			case ADSTYPE_PRINTABLE_STRING
				strOutput = strOutput & vbtab & v.PrintableString &vbcrlf
			case ADSTYPE_NUMERIC_STRING
				strOutput = strOutput & vbtab & v.NumericString &vbcrlf
			case ADSTYPE_BOOLEAN
				strOutput = strOutput & vbtab & v.Boolean & vbcrlf
			case ADSTYPE_INTEGER
				strOutput = strOutput & vbtab & v.Integer & vbcrlf	
			case ADSTYPE_OCTET_STRING
				strOutput = strOutput & vbtab & "[ADSTYPE_OCTET_STRING]" & vbcrlf
			case ADSTYPE_UTC_TIME
				strOutput = strOutput & vbtab & v.UTCTime & vbcrlf
			case ADSTYPE_LARGE_INTEGER
				strOutput = strOutput & vbtab & _
				      (v.LargeInteger.HighPart *2^32 + v.LargeInteger.LowPart) & vbcrlf
			case ADSTYPE_PROV_SPECIFIC 
				strOutput = strOutput & vbtab & "[ADSTYPE_PROV_SPECIFIC]" & vbcrlf
			case ADSTYPE_OBJECT_CLASS
				strOutput = strOutput & vbtab & "[ADSTYPE_OBJECT_CLASS]" & vbcrlf
			case ADSTYPE_CASEIGNORE_LIST
				strOutput = strOutput & vbtab & "[ADSTYPE_CASEIGNORE_LIST]" & vbcrlf
			case ADSTYPE_PATH
				strOutput = strOutput & vbtab & "[ADSTYPE_PATH]" & vbcrlf
			case ADSTYPE_POSTALADDRESS
				strOutput = strOutput & vbtab & "[ADSTYPE_POSTALADDRESS]" & vbcrlf
			case ADSTYPE_TIMESTAMP
				strOutput = strOutput & vbtab & "[ADSTYPE_TIMESTAMP]" & vbcrlf
			case ADSTYPE_BACKLINK
				strOutput = strOutput & vbtab & "[ADSTYPE_BACKLINK]" & vbcrlf
			case ADSTYPE_TYPEDNAME
				strOutput = strOutput & vbtab & "[ADSTYPE_TYPEDNAME]" & vbcrlf
			case ADSTYPE_HOLD
				strOutput = strOutput & vbtab & "[ADSTYPE_HOLD]" & vbcrlf
			case ADSTYPE_NETADDRESS
				strOutput = strOutput & vbtab & "[ADSTYPE_NETADDRESS]" & vbcrlf
			case ADSTYPE_REPLICAPOINTER
				strOutput = strOutput & vbtab & "[ADSTYPE_REPLICAPOINTER]" & vbcrlf
			case ADSTYPE_FAXNUMBER
				strOutput = strOutput & vbtab & "[ADSTYPE_FAXNUMBER]" & vbcrlf
			case ADSTYPE_EMAIL
				strOutput = strOutput & vbtab & "[ADSTYPE_EMAIL]" & vbcrlf
			case ADSTYPE_NT_SECURITY_DESCRIPTOR
				strOutput = strOutput & vbtab & "[ADSTYPE_NT_SECURITY_DESCRIPTOR]" & vbcrlf
			case ADSTYPE_UNKNOWN
				strOutput = strOutput & vbtab & "[ADSTYPE_UNKNOWN]" & vbcrlf

		end select
	next
	
next 

wscript.echo strOutput