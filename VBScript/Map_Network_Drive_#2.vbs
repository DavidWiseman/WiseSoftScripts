'**** Map a network drive ****
MapDrive "W:", "\\server\share"


'**** Function to map network drive. ****
'**** Other drive letters are tried if the specified drive letter is in use ****
sub MapDrive(byval driveLetter, byval uncPath)
	on error resume next
	set objNetwork = createobject("Wscript.Network")
	objNetwork.MapNetworkDrive driveLetter, uncPath
	
	if err.number = -2147024811 then '**** Device name in use ****
		err.clear
		driveletter = getNextDrive
		if driveLetter <> "" then
			MapDrive driveLetter, uncPath
		end if

	elseif err.number <> 0 then '**** Handle other errors ****
		wscript.echo "Unable to map network drive: " & err.description
		err.clear
	end if
end sub

'**** Function to return the first available drive letter ****
function getNextDrive 
	set objFSO = createobject("Scripting.FileSystemObject")

	for i = 68 to 88 '** Start at ASCI number 68 (D)
		drive = chr(i) & ":"
		if not objFSO.DriveExists(drive) then
			getNextDrive = drive
			exit function
		end if
	next
	wscript.echo "No more available drive letters!"

end function