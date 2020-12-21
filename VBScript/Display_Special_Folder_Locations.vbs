set objShell = createobject("Wscript.Shell")
Set objFso = CreateObject("Scripting.FileSystemObject")

with objShell
	specialFolders = "Desktop: " & .SpecialFolders("Desktop") & vbcrlf & _
			 "MyDocuments: " & .SpecialFolders("MyDocuments") & vbcrlf & _
			 "Startmenu: " & .SpecialFolders("startmenu") & vbcrlf & _
			 "AllUsersDesktop: " & .SpecialFolders("AllUsersDesktop") & vbcrlf & _
			 "AllUsersStartMenu: " & .SpecialFolders("AllUsersStartMenu") & vbcrlf & _
			 "AllUsersPrograms: " & .SpecialFolders("AllUsersPrograms") & vbcrlf & _
			 "AllUsersStartup: " & .SpecialFolders("AllUsersStartup") & vbcrlf & _
			 "Favorites: " & .SpecialFolders("Favorites") & vbcrlf & _
			 "Fonts:" & .SpecialFolders("Fonts") & vbcrlf & _
			 "NetHood: " & .SpecialFolders("NetHood") & vbcrlf & _
			 "PrintHood: " & .SpecialFolders("PrintHood") & vbcrlf & _
			 "Programs: " & .SpecialFolders("Programs") & vbcrlf & _
			 "Recent: " & .SpecialFolders("Recent") & vbcrlf & _
			 "SendTo: " & .SpecialFolders("SendTo") & vbcrlf & _
			 "Startup: " & .SpecialFolders("Startup") & vbcrlf & _
			 "Templates: " & .SpecialFolders("Templates")	
end with

'***** Other Special Folders *****

specialFolders = SpecialFolders & vbcrlf & _
		 "Windows: " &  objFSO.GetSpecialFolder(0) & vbcrlf & _
		 "System32: " & objFSO.GetSpecialFolder(1) & vbcrlf & _
		 "Temp:" & objFSO.GetSpecialFolder(2)


wscript.echo specialFolders