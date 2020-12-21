strComputer = "." ' Local Computer
strUser = "User01"

Set objUser = GetObject("WinNT://" & strComputer & "/" & strUser)
objUser.AccountDisabled = True
objUser.SetInfo