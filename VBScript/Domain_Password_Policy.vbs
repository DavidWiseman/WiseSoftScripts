set objRoot = getobject("LDAP://RootDSE")
set objDomain = getobject("LDAP://" & objRoot.get("defaultNamingContext"))

maximumPasswordAge = int(Int8ToSec(objDomain.get("maxPwdAge")) / 86400) 'convert to days
minimumPasswordAge = Int8ToSec(objDomain.get("minPwdAge")) / 86400  'convert to days
minimumPasswordLength = objDomain.get("minPwdLength")
accountLockoutDuration = Int8ToSec(objDomain.get("lockoutDuration")) / 60  'convert to minutes
lockoutThreshold = objDomain.get("lockoutThreshold") 
lockoutObservationWindow = Int8ToSec(objDomain.get("lockoutObservationWindow")) / 60 'convert to minutes
passwordHistory = objDomain.get("pwdHistoryLength")

wscript.echo "Maximum Password Age: " & maximumPasswordAge & " days" & vbcrlf & _
	     "Minimum Password Age: " & minimumPasswordAge & " days" & vbcrlf & _
	     "Enforce Password History: " & passwordHistory & " passwords remembered" & vbcrlf & _
	     "Minimum Password Length: " & minimumPasswordLength & " characters" & vbcrlf & _
	     "Account Lockout Duration: " & accountLockoutDuration & " minutes" & vbcrlf & _
	     "Account Lockout Threshold: " & lockoutThreshold & " invalid logon attempts" & vbcrlf & _
	     "Reset account lockout counter after: " & lockoutObservationWindow & " minutes"
	     

Function Int8ToSec(ByVal objInt8)
        ' Function to convert Integer8 attributes from
        ' 64-bit numbers to seconds.
        Dim lngHigh, lngLow
        lngHigh = objInt8.HighPart
        ' Account for error in IADsLargeInteger property methods.
        lngLow = objInt8.LowPart
        If lngLow < 0 Then
            lngHigh = lngHigh + 1
        End If
        Int8ToSec = -(lngHigh * (2 ^ 32) + lngLow) / (10000000)
End Function