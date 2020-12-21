' Local Computer
strComputer = "."

Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")
Set colItems = objWMIService.ExecQuery("Select * From Win32_WinSAT")

For Each objItem in colItems
    Wscript.Echo "Processor: " & objItem.CPUScore
    Wscript.Echo "Memory: " & objItem.MemoryScore
    Wscript.Echo "Primary hard disk: " & objItem.DiskScore
    Wscript.Echo "Graphics: " & objItem.GraphicsScore
    Wscript.Echo "Gaming graphics: " & objItem.D3DScore
    Wscript.Echo "Windows System Performance Rating: " & objItem.WinCRSLevel
Next
