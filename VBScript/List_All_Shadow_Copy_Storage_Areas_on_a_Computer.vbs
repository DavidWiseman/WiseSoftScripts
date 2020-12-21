strComputer = "."
Set objWMIService = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")

Set colItems = objWMIService.ExecQuery("Select * from Win32_ShadowStorage")

For Each objItem in colItems
    Wscript.Echo "Volume: " & objItem.Volume
    Wscript.Echo "Allocated space: " & objItem.AllocatedSpace
    Wscript.Echo "Differential volume: " & objItem.DiffVolume
    Wscript.Echo "Maximum space: " & objItem.MaxSpace
    Wscript.Echo "Used space: " & objItem.UsedSpace
    Wscript.Echo
Next
