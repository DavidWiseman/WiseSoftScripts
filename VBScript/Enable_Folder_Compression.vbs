strFolder = "C:\Test"
strComputer = "."

Set objWMIService = GetObject("winmgmts:" _
		 & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")

strFolder = REPLACE(strFolder,"\","\\")

Set colFolders = objWMIService.ExecQuery("Select * from Win32_Directory where name = '" & strFolder & "'")

For Each objFolder in colFolders
   	result = objFolder.Compress
	if result <> 0 then
		wscript.echo "Unable to compress folder:" & objFolder.Name
	end if
Next	
