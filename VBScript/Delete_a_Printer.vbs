strComputer = "."
Set objWMIService = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")

Set colInstalledPrinters =  objWMIService.ExecQuery _
    ("Select * from Win32_Printer where DeviceID = 'ScriptedPrinter'")

For Each objPrinter in colInstalledPrinters
    objPrinter.Delete_
Next
