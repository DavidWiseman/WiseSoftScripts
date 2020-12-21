strComputer = "." ' Local Computer

Set objComputer = GetObject("WinNT://" & strComputer)
objComputer.Filter = Array("user")

For Each objUser In objComputer
	Wscript.Echo objUser.Name 
Next