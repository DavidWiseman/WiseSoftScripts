Set objSysInfo = CreateObject("ADSystemInfo")

Wscript.Echo "User name: " & objSysInfo.UserName & vbcrlf & _
 	     "Computer name: " & objSysInfo.ComputerName & vbcrlf & _
 	     "Site name: " & objSysInfo.SiteName & vbcrlf & _
	     "Domain short name: " & objSysInfo.DomainShortName & vbcrlf & _
	     "Domain DNS name: " & objSysInfo.DomainDNSName & vbcrlf & _
	     "Forest DNS name: " & objSysInfo.ForestDNSName & vbcrlf & _
             "PDC role owner: " & objSysInfo.PDCRoleOwner & vbcrlf & _
             "Schema role owner: " & objSysInfo.SchemaRoleOwner & vbcrlf & _
             "Domain is in native mode: " & objSysInfo.IsNativeMode