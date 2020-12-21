On Error Resume Next
   
Set objServer = GetObject _
    ("LDAP://CN=wisesoft-DC01,CN=Servers,CN=Default-First-Site-Name,"  & _
        " CN=Sites,CN=Configuration,DC=wisesoft,DC=co,DC=uk")
 
dnBHTList = objServer.GetEx("bridgeheadTransportList")
 
WScript.Echo "Bridge Head Transport List:"
WScript.Echo "This multi-valued attribute lists the protocol" & _
    "transports over which this BridgeHead Server replicates"
For Each dnValue in dnBHTList
    WScript.Echo "Value: " & dnValue
Next
