strComputer = "."
Set objWMIService = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")

Set colItems = objWMIService.ExecQuery("Select * from Win32_ShadowContext")

For Each objItem in colItems
    Wscript.Echo "Name: " & objItem.Name
    Wscript.Echo "Client accessible: " & objItem.ClientAccessible
    Wscript.Echo "Differential: " & objItem.Differential
    Wscript.Echo "Exposed locally: " & objItem.ExposedLocally
    Wscript.Echo "Exposed remotely: " & objItem.ExposedRemotely
    Wscript.Echo "Hardware assisted: " & objItem.HardwareAssisted
    Wscript.Echo "Imported: " & objItem.Imported
    Wscript.Echo "No auto release: " & objItem.NoAutoRelease
    Wscript.Echo "Not surfaced: " & objItem.NotSurfaced
    Wscript.Echo "No writers: " & objItem.NoWriters
    Wscript.Echo "Persistent: " & objItem.Persistent
    Wscript.Echo "Plex: " & objItem.Plex
    Wscript.Echo "Transportable: " & objItem.Transportable
    Wscript.Echo
Next
