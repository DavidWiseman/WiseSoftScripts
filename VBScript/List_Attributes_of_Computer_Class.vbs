Set objSchema = GetObject("LDAP://schema/computer")
 
strOutput = "Mandatory attributes:" & vbcrlf & vbcrlf

For Each strAttribute in objSchema.MandatoryProperties
	strOutput = strOutput & vbtab & strAttribute & vbcrlf
Next
 
strOutput = strOutput & vbcrlf & "Optional Attributes:" & vbcrlf & vbcrlf

For Each strAttribute in objSchema.OptionalProperties
	strOutput = strOutput & vbtab & strAttribute & vbcrlf
Next
	
wscript.echo strOutput