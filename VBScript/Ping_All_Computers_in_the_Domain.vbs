OPTION Explicit
DIM cn,cmd,rs
DIM objRoot
DIM intFailed, intSucceeded
DIM strPing

set cmd = createobject("ADODB.Command")
set cn = createobject("ADODB.Connection")
set rs = createobject("ADODB.Recordset")

cn.open "Provider=ADsDSOObject;"
cmd.activeconnection = cn

' Used to get the default naming context. e.g. dc=wisesoft,dc=co,dc=uk
set objRoot = getobject("LDAP://RootDSE")

' Query for all computers in the domain
cmd.commandtext = "<LDAP://" & objRoot.get("defaultNamingContext") & ">;(objectCategory=Computer);" & _
		  "dnsHostName;subtree"
'**** Bypass 1000 record limitation ****
cmd.properties("page size")=1000

set rs = cmd.execute

intFailed = 0
intSucceeded = 0

' Ping all computers in the domain
while rs.eof <> true and rs.bof <> true
	strPing = ping(rs("dnsHostName"))
	IF LEFT(strPing,2) = "OK" then
		intSucceeded = intSucceeded + 1
	ELSE
		intFailed = intFailed + 1
	END IF

	wscript.echo rs("dnsHostName") & " : " & strPing
	rs.movenext
wend

cn.close

wscript.echo "Finished (" & intSucceeded & " Succeeded, " & intFailed & " Failed)"

' Function to ping a computer
private function ping(byval strComputer)
	DIM Status,objPing, ObjPingStatus
	status = "Error"
	Set objPing = GetObject("winmgmts:{impersonationLevel=impersonate}")._
        		ExecQuery("select * from Win32_PingStatus where address = '" & _
            		strComputer & "'")
	For Each objPingStatus in objPing

        	If IsNull(objPingStatus.StatusCode) then
			status = "Failed"
		elseif objPingStatus.StatusCode<>0 Then 
			status = "Failed (" & getPingStatus(objPingStatus.StatusCode) & ")"
		else
			status = "OK (Bytes= " & objPingStatus.BufferSize & _
				", Time = " & objPingStatus.ResponseTime & _
				", TTL = " & objPingStatus.ResponseTimeToLive & ")"
        	End If
    	Next

	ping = status
end function

' Function to convert the status code into a useful description
private function getPingStatus(byval statusCode)
	DIM status
	status = statusCode
	SELECT CASE statusCode
		CASE 11001
 			status = "Buffer Too Small"
 		CASE 11002
			status = "Destination Net Unreachable"
 		CASE 11003
 			status = "Destination Host Unreachable"
 		CASE 11004
 			status = "Destination Protocol Unreachable"
 		CASE 11005
 			status = "Destination Port Unreachable"
 		CASE 11006
 			status = "No Resources"
 		CASE 11007
 			status = "Bad Option"
 		CASE 11008
			status = "Hardware Error"
 		CASE 11009
 			status = "Packet Too Big"
 		CASE 11010
 			status = "Request Timed Out"
 		CASE 11011
 			status = "Bad Request"
 		CASE 11012
			status = "Bad Route"
 		CASE 11013
 			status = "TimeToLive Expired Transit"
 		CASE 11014
 			status = "TimeToLive Expired Reassembly"
 		CASE 11015
 			status = "Parameter Problem"
 		CASE 11016
			status = "Source Quench"
 		CASE 11017
 			status = "Option Too Big"
 		CASE 11018
			status = "Bad Destination"
 		CASE 11032
			status = "Negotiating IPSEC"
 		CASE 11050
			status = "General Failure"
	END SELECT
	getPingStatus = status
end function
 