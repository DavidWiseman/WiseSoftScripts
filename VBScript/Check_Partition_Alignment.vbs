strComputer = "." ' Local Computer
Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\CIMV2") 
Set colItems = objWMIService.ExecQuery("SELECT * FROM Win32_DiskPartition",,48) 

For Each objItem in colItems     
	Wscript.Echo "Disk: " & objItem.DiskIndex & "  Partition: " & objItem.Index & "  StartingOffset: " & objItem.StartingOffset/1024 & "KB"     
Next