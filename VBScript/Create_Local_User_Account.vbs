strComputer = "." ' Local Computer
strUser = "User01"
strPassword = "P@$$W0rd"

Set colAccounts = GetObject("WinNT://" & strComputer & "")
Set objUser = colAccounts.Create("user", strUser)
objUser.SetPassword strPassword
objUser.SetInfo