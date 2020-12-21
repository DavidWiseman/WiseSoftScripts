Const LocalDocumentsFolder = "C:\Documents and Settings\"

set objFSO = createobject("Scripting.FileSystemObject")
set objFolder = objFSO.GetFolder(localdocumentsfolder)

on error resume next

for each fldr in objFolder.SubFolders
	if not isexception(fldr.name) then
		objFSO.DeleteFolder fldr.path, True
	end if
next


Function isException(byval foldername)
	select case foldername
		case "All Users"
			isException = True
		case "Default User"
			isException = True
		case "LocalService"
			isException = True
		case "NetworkService"
			isException = True
		case "Administrator"
			isException = True
		case Else
			isException = False
	End Select
End Function