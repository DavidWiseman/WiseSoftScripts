strDcName = "wisesoft-dc01"
strSiteName = "my-site-name"
 
Set objADSysInfo = CreateObject("ADSystemInfo")
strDcSiteName = objADSysInfo.GetDCSiteName(strDcName)
 
If UCase(strSiteName) = UCase(strDcSiteName) Then
    WScript.Echo "TRUE: " & strDcName & " is in site " & strSiteName
Else
    WScript.Echo "FALSE: " & strDcName & " is NOT in site " & strSiteName
End If
