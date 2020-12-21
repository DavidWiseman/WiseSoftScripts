Set objRootDSE = GetObject("LDAP://rootDSE")

' Schema Master
Set objSchema = GetObject _
    ("LDAP://" & objRootDSE.Get("schemaNamingContext"))
strSchemaMaster = objSchema.Get("fSMORoleOwner")
Set objNtds = GetObject("LDAP://" & strSchemaMaster)
Set objComputer = GetObject(objNtds.Parent)
strSchemaMaster = objComputer.dNSHostName

' Domain Naming Master
Set objPartitions = GetObject("LDAP://CN=Partitions," & _ 
    objRootDSE.Get("configurationNamingContext"))
strDomainNamingMaster = objPartitions.Get("fSMORoleOwner")
Set objNtds = GetObject("LDAP://" & strDomainNamingMaster)
Set objComputer = GetObject(objNtds.Parent)
strDomainNamingMaster = objComputer.dNSHostName

' PDC Emulator
Set objDomain = GetObject _
    ("LDAP://" & objRootDSE.Get("defaultNamingContext"))
strPdcEmulator = objDomain.Get("fSMORoleOwner")
Set objNtds = GetObject("LDAP://" & strPdcEmulator)
Set objComputer = GetObject(objNtds.Parent)
strPdcEmulator = objComputer.dNSHostName

' RID Master
Set objRidManager = GetObject("LDAP://CN=RID Manager$,CN=System," & _
    objRootDSE.Get("defaultNamingContext"))
strRidMaster = objRidManager.Get("fSMORoleOwner")
Set objNtds = GetObject("LDAP://" & strRidMaster)
Set objComputer = GetObject(objNtds.Parent)
strRidMaster = objComputer.dNSHostName

' Infrastructure Master
Set objInfrastructure = GetObject("LDAP://CN=Infrastructure," & _
    objRootDSE.Get("defaultNamingContext"))
strInfrastructureMaster = objInfrastructure.Get("fSMORoleOwner")
Set objNtds = GetObject("LDAP://" & strInfrastructureMaster)
Set objComputer = GetObject(objNtds.Parent)
strInfrastructureMaster = objComputer.dNSHostName


WScript.Echo "Forest-wide Domain Naming Master FSMO: " & strDomainNamingMaster & vbcrlf & _
	     "Forest-wide Schema Master FSMO: " & strSchemaMaster & vbcrlf & _
	     "Domain's Infrastructure Master FSMO: " & strInfrastructureMaster & vbcrlf & _
	     "Domain's RID Master FSMO: " & strRidMaster & vbcrlf & _
	     "Domain's PDC Emulator FSMO: " & strPdcEmulator