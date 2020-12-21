strComputer = "." ' Local Computer
strUser = "User01"

Set objUser = GetObject("WinNT://" & strComputer & "/" & strUser)
objUser.Put "PasswordExpired", 1
objUser.SetInfo