Set objRoot = GetObject("LDAP://rootDSE")
strDC = objRoot.Get("dnsHostName")
Wscript.Echo "Authenticated By: " & strDC