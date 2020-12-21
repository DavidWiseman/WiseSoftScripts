strComputer = "."

Set objWMIService = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")

Set colDisks = objWMIService.ExecQuery("Select * from Win32_LogicalDisk")

For Each objDisk in colDisks
    Wscript.Echo objDisk.DeviceID & vbTab & objDisk.FileSystem
Next