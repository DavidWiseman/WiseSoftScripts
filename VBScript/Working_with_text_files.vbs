'<------------------How to write to a new text File--------------------------->


textFilePath = "C:\wisesoft_test.txt"
set objFSO = createobject("Scripting.FileSystemObject")

set objTextFile = objFSO.CreateTextFile(textFilePath)

objTextFile.WriteLine("This file was created " & now())
objTextFile.Write("The WriteLine method inserts a new line automatically at the end of the text and the write method ")
objTextFile.WriteLine("allows you to continue on the same line.")
objTextFile.WriteLine("Visit http://www.wisesoft.co.uk for more script samples.")

objTextFile.Close


'<------------------How to append data to a text file------------------------->


const ForAppending = 8
textFilePath = "C:\wisesoft_test.txt"
set objFSO = createobject("Scripting.FileSystemObject")
set objTextFile = objFSO.opentextfile(textFilePath,ForAppending)

objTextFile.WriteLine("I added this text later")

objTextFile.Close


'<------------------How to read from a text File ----------------------------->


textFilePath = "C:\wisesoft_test.txt"
set objFSO = createobject("Scripting.FileSystemObject")
set objTextFile = objFSO.opentextfile(textFilePath)

wscript.echo objTextFile.ReadAll

objTextFile.Close


'<------------------How to read to a text File line by line ------------------>


textFilePath = "C:\wisesoft_test.txt"
set objFSO = createobject("Scripting.FileSystemObject")
set objTextFile = objFSO.opentextfile(textFilePath)

do until objTextFile.AtEndOfStream
	wscript.echo objTextFile.ReadLine
loop

objTextFile.Close