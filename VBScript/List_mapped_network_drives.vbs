const bytesToGb = 1073741824
strComputer = "." ' Local Computer
set objWMI = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")
set colDrives = objWMI.ExecQuery("select * from Win32_MappedLogicalDisk")

for each objDrive in colDrives
	WScript.Echo "Device ID: " & objDrive.DeviceID & vbcrlf & _
   		     "Volume Name: " & objDrive.VolumeName & vbcrlf & _
		     "Session ID: " & objDrive.SessionID & vbcrlf & _
		     "Size: " & round(objDrive.Size / bytesToGb,1) & " Gb"
next
