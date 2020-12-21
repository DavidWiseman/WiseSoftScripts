strFolder = "C:\Windows"
set objFSO = createobject("Scripting.FileSystemObject")

if objFSO.FolderExists(strFolder) then
	wscript.echo "The folder exists"
else
	wscript.echo "Sorry, this folder does not exist"
end if