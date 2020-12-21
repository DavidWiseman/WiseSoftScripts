strFile= "C:\Windows\notepad.exe"
set objFSO = createobject("Scripting.FileSystemObject")

if objFSO.FileExists(strFile) then
	wscript.echo "The file exists"
else
	wscript.echo "Sorry, the file does not exist"
end if