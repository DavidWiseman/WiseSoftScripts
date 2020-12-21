OPTION EXPLICIT
DIM strExtensionsToDelete,strFolder
DIM objFSO

' ************************************************************
' Setup
' ************************************************************

' Folder to delete files from (files will also be deleted from subfolders)
strFolder = "R:\PublishTemp"
' A comma separated list of file extensions
' Files with extensions provided in the list below will be deleted
strExtensionsToDelete = "wav,avi,mp3"

' ************************************************************

set objFSO = createobject("Scripting.FileSystemObject")

RecursiveDeleteByExtension strFolder,strExtensionsToDelete

wscript.echo "Finished"

sub RecursiveDeleteByExtension(byval strDirectory,strExtensionsToDelete)
	DIM objFolder, objSubFolder, objFile
	DIM strExt

	set objFolder = objFSO.GetFolder(strDirectory)
	for each objFile in objFolder.Files
		for each strExt in SPLIT(UCASE(strExtensionsToDelete),",")
			if RIGHT(UCASE(objFile.Path),LEN(strExt)+1) = "." & strExt then
				wscript.echo "Deleting:" & objFile.Path
				objFile.Delete
				exit for
			end if
		next
	next	
	for each objSubFolder in objFolder.SubFolders
		RecursiveDeleteByExtension objSubFolder.Path,strExtensionsToDelete
	next
end sub