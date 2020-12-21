Option Explicit

const strComputer = "."
const strReport = "c:\diskspace.txt"


Dim objWMIService, objItem, colItems
Dim strDriveType, strDiskSize, txt

Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")
Set colItems = objWMIService.ExecQuery("Select * from Win32_LogicalDisk WHERE DriveType=3")
txt = "Drive" & vbtab & "Size" & vbtab & "Used" & vbtab & "Free" & vbtab & "Free(%)" & vbcrlf
For Each objItem in colItems
	
	DIM pctFreeSpace,strFreeSpace,strusedSpace
	
	pctFreeSpace = INT((objItem.FreeSpace / objItem.Size) * 1000)/10
	strDiskSize = Int(objItem.Size /1073741824) & "Gb"
	strFreeSpace = Int(objItem.FreeSpace /1073741824) & "Gb"
	strUsedSpace = Int((objItem.Size-objItem.FreeSpace)/1073741824) & "Gb"
	txt = txt & objItem.Name & vbtab & strDiskSize & vbtab & strUsedSpace & vbTab & strFreeSpace & vbtab & pctFreeSpace & vbcrlf

Next

writeTextFile txt, strReport
wscript.echo "Report written to " & strReport & vbcrlf & vbcrlf & txt

' Procedure to write output to a text file
private sub writeTextFile(byval txt,byval strTextFilePath)
	Dim objFSO,objTextFile
	
	set objFSO = createobject("Scripting.FileSystemObject")

	set objTextFile = objFSO.CreateTextFile(strTextFilePath)

	objTextFile.Write(txt)

	objTextFile.Close
	SET objTextFile = nothing
end sub
