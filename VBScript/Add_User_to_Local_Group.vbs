' Change Parameters as required
strComputer = "computer01"
strGroup = "Administrators"
strUser = "WiseSoftUser01"

' Get group object
Set objGroup = GetObject("WinNT://" & strComputer & "/" & strGroup & ",group")
' Get user object
Set objUser = GetObject("WinNT://" & strComputer & "/" & strUser & ",user")
' Add user to group
objGroup.Add(objUser.ADsPath)