const MaxConnections = 0

set objFSO = createobject("Scripting.FileSystemObject")

RootFolder = inputbox("Please enter the root folder that contains the folders you want to share:")

'***** Perform some basic validation *****
if not objFSO.FolderExists(RootFolder) then
	wscript.echo "Invalid Folder"
	wscript.quit
end if
if mid(RootFolder,1,2)="\\" then
	wscript.echo "UNC Paths are not supported"
	wscript.quit
end if

set objRootFolder = objFSO.GetFolder(rootfolder)
for each fldr in objRootFolder.SubFolders
	sharefolder fldr.path, fldr.name
next

wscript.echo "Completed"

sub shareFolder(byval folderPath,shareName)
	strComputer = "."
	Set objWMIService = GetObject("winmgmts:{impersonationLevel=impersonate}!\\" & _
							strComputer & "\root\cimv2")
	Set objNewShare = objWMIService.Get("Win32_Share")

	objNewShare.Create folderpath, shareName, MaxConnections
end sub
