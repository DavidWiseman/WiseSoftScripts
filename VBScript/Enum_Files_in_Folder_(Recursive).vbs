strFolder = "C:\test"

set objFSO = createobject("Scripting.FileSystemObject")

GetFiles strFolder

sub GetFiles(byval strDirectory)
	set objFolder = objFSO.GetFolder(strDirectory)
	for each objFile in objFolder.Files
		wscript.echo objFile.Path 
	next	
	for each objFolder in objFolder.SubFolders
		GetFiles objFolder.Path
	next
end sub