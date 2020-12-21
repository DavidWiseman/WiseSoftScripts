strFolder = "C:\test"

set objFSO = createobject("Scripting.FileSystemObject")

GetFolders strFolder

sub GetFolders(byval strDirectory)
	set objFolder = objFSO.GetFolder(strDirectory)	
	for each objFolder in objFolder.SubFolders
		wscript.echo objFolder.Path
		GetFolders objFolder.Path
	next
end sub