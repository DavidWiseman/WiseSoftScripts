strComputer = "."
Set objWMIService = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")
)
Set colItems = objWMIService.ExecQuery("Select * From Win32_ShadowCopy")

For Each objItem in colItems
    objItem.Delete_
Next
