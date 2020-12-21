strComputer = "." ' Local Computer
strUser = "User01"

Set objUser = GetObject("WinNT://" & strComputer & "/" & strUser)
objUser.AccountExpirationDate = #01/01/2010# 
objUser.SetInfo