strSource = "c:\myfile.txt"
strDestination = "C:\newfolder\myfile.txt"


set objFSO = createobject("Scripting.FileSystemObject")
objFSO.CopyFile strSource,strDestination