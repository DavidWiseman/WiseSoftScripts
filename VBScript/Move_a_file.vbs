strSource = "c:\myfile.txt"
strDestination = "C:\newfolder\myfile.txt"

set objFSO = createobject("Scripting.FileSystemObject")
objFSO.MoveFile strSource,strDestination