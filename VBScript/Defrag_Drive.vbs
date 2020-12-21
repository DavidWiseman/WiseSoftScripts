strComputer = "." ' Local Computer
strDrive = "C:"  ' Specify Drive
boolForce = False ' Force defragmentation, even if freespace is low

' Connect to WMI
set objWMI = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")
' Run WMI query to return a colection of volumes
set colVol = objWMI.ExecQuery("select * from Win32_Volume Where Name = '" & _
                              strDrive & "\\'")
' For each volume in collection
for each objVol in colVol
	' Run Defrag Method
	intRC = objVol.Defrag(boolForce,objRpt)
	' Display message when defrag completes
	select case intRC
	case 0
		WScript.Echo "Finished defragmenting drive " &  objVol.DriveLetter
	case 1
		wscript.echo "Error defregmenting drive " & objVol.DriveLetter & _
			     vbcrlf & "1 Access Denied"
   	case 2
		wscript.echo "Error defregmenting drive " & objVol.DriveLetter & _
			     vbcrlf & "2 Not supported"
	case 3
		wscript.echo "Error defregmenting drive " & objVol.DriveLetter & _
			      vbcrlf & "3 Volume dirty bit is set"
	case 4
		wscript.echo "Error defregmenting drive " & objVol.DriveLetter & _
			     vbcrlf & "4 Not enough free space"
	case 5
		wscript.echo "Error defregmenting drive " & objVol.DriveLetter & _
			     vbcrlf & "5 Corrupt Master File Table detected"
	case 6
		wscript.echo "Error defregmenting drive " & objVol.DriveLetter & _
			     vbcrlf & "6 Call canceled"
	case 7
		wscript.echo "Error defregmenting drive " & objVol.DriveLetter & _
			     vbcrlf & "7 Call cancellation request too late"
	case 8
		wscript.echo "Error defregmenting drive " & objVol.DriveLetter & _
			     vbcrlf & "8 Defrag engine is already running"
	case 9
		wscript.echo "Error defregmenting drive " & objVol.DriveLetter & _
			     vbcrlf & "9 Unable to connect to defrag engine"
	case 10
		wscript.echo "Error defregmenting drive " & objVol.DriveLetter & _
			     vbcrlf & "10 Defrag engine error"	
	case else '11 Unknown Error / Other

		wscript.echo "Error defregmenting drive " & objVol.DriveLetter & _
			     vbcrlf & intRC & " Unknown error"
	end select
next