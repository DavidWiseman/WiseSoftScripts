' Change Parameters as required
strComputer ="computer01"
strGroup = "Administrators"
strUser = "test"

Set objGroup = GetObject("WinNT://" & strComputer & "/" & strGroup & ",group")
Set objUser = GetObject("WinNT://" & strComputer & "/" & strUser & ",user")
 
objGroup.Remove(objUser.ADsPath)