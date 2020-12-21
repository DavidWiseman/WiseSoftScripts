strComputer = "."
Set objWMIService = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")

Set colPrinters =  objWMIService.ExecQuery _
    ("Select * from Win32_Printer Where DeviceID = 'HP LaserJet 4Si M'")

For Each objPrinter in colPrinters
    objPrinter.RenamePrinter("ArtDepartmentPrinter")
Next

Set colPrinters = objWMIService.ExecQuery _
    ("Select * From Win32_Printer Where DeviceID = 'ArtDepartmentPrinter' ")

For Each objPrinter in colPrinters
    objPrinter.ShareName = "ArtDepartmentPrinter"
    objPrinter.Put_
Next
