' Change Parameters as required
strComputer ="computer01"
strGroup = "WiseSoftGroup"

Set objComputer = GetObject("WinNT://" & strComputer & "")
objComputer.Delete "group", strGroup