OPTION EXPLICIT
DIM strExtensionsToDelete,strFolder
DIM objFSO, MaxAge, IncludeSubFolders

' ************************************************************
' Setup
' ************************************************************

' Folder to delete files
strFolder = "C:\test\"
' Delete files from sub-folders?
includeSubfolders = true
' A comma separated list of file extensions
' Files with extensions provided in the list below will be deleted
strExtensionsToDelete = "tmp,temp"
' Max File Age (in Days).  Files older than this will be deleted.
maxAge = 10

' ************************************************************

set objFSO = createobject("Scripting.FileSystemObject")

DeleteFiles strFolder,strExtensionsToDelete, maxAge, includeSubFolders

wscript.echo "Finished"

sub DeleteFiles(byval strDirectory,byval strExtensionsToDelete,byval maxAge,includeSubFolders)
	DIM objFolder, objSubFolder, objFile
	DIM strExt

	set objFolder = objFSO.GetFolder(strDirectory)
	for each objFile in objFolder.Files
		for each strExt in SPLIT(UCASE(strExtensionsToDelete),",")
			if RIGHT(UCASE(objFile.Path),LEN(strExt)+1) = "." & strExt then
				IF objFile.DateLastModified < (Now - MaxAge) THEN
					wscript.echo "Deleting:" & objFile.Path & " | " & objFile.DateLastModified 
					objFile.Delete
					exit for
				END IF
			end if
		next
	next	
	if includeSubFolders = true then ' Recursive delete
		for each objSubFolder in objFolder.SubFolders
			DeleteFiles objSubFolder.Path,strExtensionsToDelete,maxAge, includeSubFolders
		next
	end if
end sub