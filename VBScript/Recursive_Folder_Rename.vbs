Root = inputbox("Please enter the root folder (all subfolders will be renamed)" & vbcrlf & "e.g. C:\TEST")
if Root="" then Canceled

FindStr = inputbox("Please enter the string that you want to find")
if FindStr = "" then Canceled

ReplaceStr = inputbox("Please enter the string that you want to replace it with")
if ReplaceStr = "" then Canceled

set objFSO= createobject("Scripting.FileSystemObject")

EnumFolders Root

Sub EnumFolders(byval Folder)
	set objFolder = objFSO.GetFolder(Folder)
	set colSubfolders = objFolder.Subfolders

	for each objSubfolder in colSubfolders
		NewFolderName = (Replace(objSubfolder.name, findstr, replacestr))
			If NewFolderName <> objSubFolder.Name Then
				objSubFolder.Name = NewFolderName
			End If
		enumfolders objSubfolder.path
	next

End Sub

sub Canceled
	wscript.echo "Script Canceled"
	wscript.quit
end sub