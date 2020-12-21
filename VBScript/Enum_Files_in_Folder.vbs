strFolder = "C:\test"

set objFSO = createobject("Scripting.FileSystemObject")
set objFolder = objFSO.GetFolder(strFolder)

for each objFile in objFolder.Files
	wscript.echo objFile.Path 
	'wscript.echo objFile.Name
next