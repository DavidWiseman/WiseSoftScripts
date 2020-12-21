strSiteRDN = "cn=Default-First-Site-Name"
 
Set objRootDSE = GetObject("LDAP://RootDSE")
strConfigurationNC = objRootDSE.Get("configurationNamingContext")
 
strServersPath = "LDAP://cn=Servers," & strSiteRDN & ",cn=Sites," & _
    		strConfigurationNC
Set objServersContainer = GetObject(strServersPath)
 
For Each objServer In objServersContainer
	wscript.echo objServer.dNSHostName
Next