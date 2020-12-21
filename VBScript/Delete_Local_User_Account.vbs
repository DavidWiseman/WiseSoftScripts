strComputer = "." ' Local Computer
strUser = "User01"

Set objComputer = GetObject("WinNT://" & strComputer & "")
objComputer.Delete "user", strUser