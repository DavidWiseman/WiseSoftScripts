On Error Resume Next

strComputer = "."
strUser = "User01"

Set objUser = GetObject("WinNT://" & strComputer & "/" & strUser)
Set objClass = GetObject(objUser.Schema)

WScript.Echo "Mandatory properties for " & objUser.Name & ":"
For Each property In objClass.MandatoryProperties
	WScript.Echo property, objUser.Get(property)
Next

WScript.Echo "Optional properties for " & objUser.Name & ":"
For Each property In objClass.OptionalProperties
	WScript.Echo property, objUser.Get(property)
Next