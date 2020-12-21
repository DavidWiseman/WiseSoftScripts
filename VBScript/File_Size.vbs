const bytesToKb = 1024
strFile = "C:\myfile.dat"

set objFSO = createobject("Scripting.FileSystemObject")
set objFile = objFSO.GetFile(strFile)

wscript.echo "File Size: " & cint(objFile.Size / bytesToKb) & "Kb"