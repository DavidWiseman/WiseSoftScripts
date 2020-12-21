intSmallestQueue = 1000

strComputer = "."
Set objWMIService = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")

Set colPrintQueues =  objWMIService.ExecQuery _
    ("Select * from Win32_PerfFormattedData_Spooler_PrintQueue " _
        & "Where Name <> '_Total'")

For Each objPrintQueue in colPrintQueues
    intCurrentQueue = objPrintQueue.Jobs + objPrintQueue.JobsSpooling
    If intCurrentQueue <= intSmallestQueue Then
        strNewDefault = objPrintQueue.Name
        intSmallestQueue = intCurrentQueue
    End If
Next

Set colInstalledPrinters =  objWMIService.ExecQuery _
    ("Select * from Win32_Printer Where Name = '" & strNewDefault & "'")

For Each objPrinter in colInstalledPrinters
    objPrinter.SetDefaultPrinter()
Next
