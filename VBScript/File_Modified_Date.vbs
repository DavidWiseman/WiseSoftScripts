strFile = "C:\myfile.dat"

Set objFSO = CreateObject("Scripting.FileSystemObject")
set objFile = objFSO.GetFile(strFile)

wscript.echo "File Modified: " &  CDate( objFile.DateLastModified)