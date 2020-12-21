strComputers = "server1,server2,server3"
arrComputers = split(strComputers, ",")
 
For Each strComputer in arrComputers

	Set objPing = GetObject("winmgmts:{impersonationLevel=impersonate}")._
        		ExecQuery("select * from Win32_PingStatus where address = '" & _
            		strComputer & "'")
	For Each objPingStatus in objPing
        	If IsNull(objPingStatus.StatusCode) or objPingStatus.StatusCode<>0 Then 
			if strFailedPings <> "" then strFailedPings = strFailedPings & vbcrlf
			strFailedPings = strFailedPings & strComputer
        	End If
    	Next
Next

IF strFailedPings = "" then
	wscript.echo "Ping status of specified computers is OK"
ELSE
	wscript.echo "Ping failed for the following computers:" & _
			vbcrlf & vbcrlf & strFailedPings
END IF
	