strComputer = "." ' Local computer

set objWMIDateTime = CreateObject("WbemScripting.SWbemDateTime")
set objWMI = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")
set colOS = objWMI.InstancesOf("Win32_OperatingSystem")
for each objOS in colOS
	objWMIDateTime.Value = objOS.LastBootUpTime
	Wscript.Echo "Last Boot Up Time: " & objWMIDateTime.GetVarDate & vbcrlf & _
		"System Up Time: " &  TimeSpan(objWMIDateTime.GetVarDate,Now) & _
		" (hh:mm:ss)"
next

Function TimeSpan(dt1, dt2) 
	' Function to display the difference between
	' 2 dates in hh:mm:ss format
	If (isDate(dt1) And IsDate(dt2)) = false Then 
		TimeSpan = "00:00:00" 
		Exit Function 
        End If 
 
        seconds = Abs(DateDiff("S", dt1, dt2)) 
        minutes = seconds \ 60 
        hours = minutes \ 60 
        minutes = minutes mod 60 
        seconds = seconds mod 60 
 
        if len(hours) = 1 then hours = "0" & hours 
 
        TimeSpan = hours & ":" & _ 
            RIGHT("00" & minutes, 2) & ":" & _ 
            RIGHT("00" & seconds, 2) 
End Function 