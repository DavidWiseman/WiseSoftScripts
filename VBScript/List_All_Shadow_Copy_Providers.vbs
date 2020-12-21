strComputer = "."
Set objWMIService = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")

Set colItems = objWMIService.ExecQuery("Select * from Win32_ShadowProvider")

For Each objItem in colItems
    Wscript.Echo "Name: " & objItem.Name
    Wscript.Echo "CLSID: " & objItem.CLSID
    Wscript.Echo "ID: " & objItem.ID
    Wscript.Echo "Type: " & objItem.Type
    Wscript.Echo "Version: " & objItem.Version
    Wscript.Echo "Version ID: " & objItem.VersionID
    Wscript.Echo
Next
