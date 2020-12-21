' Change Parameters as required
strComputer = "computer01"
strGroup = "WiseSoftGroup"

Set objComp = GetObject("WinNT://" & strComputer & "")
Set objGroup = objComp.Create("group", strGroup)
objGroup.SetInfo
