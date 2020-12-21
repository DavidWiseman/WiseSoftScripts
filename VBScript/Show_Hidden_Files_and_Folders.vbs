'* Show Hidden Files and Folders
ShowHiddenFilesAndFolders True

'* Hide Hiden Files and Folders
ShowHiddenFilesAndFolders False


Function ShowHiddenFilesAndFolders(ByVal Enabled)

	On Error Resume Next

	set objShell = Wscript.CreateObject("WScript.Shell")
	If Enabled = True Then
        		RegValue = 1
    	Else
        		RegValue = 2
    	End If
    	RegPath = "HKCU\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced\Hidden"
    	CurrentValue = objShell.RegRead(RegPath)
    	errnum = Err.Number

    	If errnum = 0 and (RegValue<>CurrentValue) Then
        
        	objShell.Regwrite RegPath, RegValue, "REG_DWORD"
       	 	On Error GoTo 0

		'***** Close all instances of explorer so the setting is applied immediately *****
        	For Each Process In GetObject("winmgmts:"). _
                        ExecQuery("select * from Win32_Process where name='explorer.exe'")
            		Process.Terminate (0)
        	Next
   	End If

End Function