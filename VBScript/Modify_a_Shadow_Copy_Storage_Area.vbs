strComputer = "."
Set objWMIService = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")

Set colItems = objWMIService.ExecQuery("Select * from Win32_ShadowStorage")

For Each objItem in colItems
    objItem.MaxSpace = 500000000
    objItem.Put_
Next
