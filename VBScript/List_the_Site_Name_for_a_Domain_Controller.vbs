strDcName = "wisesoft-dc01"
Set objADSysInfo = CreateObject("ADSystemInfo")

strDcSiteName = objADSysInfo.GetDCSiteName(strDcName)
WScript.Echo "DC Site Name: " & strDcSiteName
