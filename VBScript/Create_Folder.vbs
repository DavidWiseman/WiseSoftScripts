strFolder = "C:\test"

set objFSO = createobject("Scripting.FileSystemObject")

if objFSO.FolderExists(strFolder) = False then
	objFSO.CreateFolder strFolder
	wscript.echo "Folder Created"
else
	wscript.echo "Folder already exists"
end if