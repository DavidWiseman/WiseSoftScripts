On Error Resume Next

Set objContainer = GetObject _
    ("LDAP://ou=Development,dc=NA,dc=wisesoft,dc=co,dc=uk") 
 
strExistingGPLink = objContainer.Get("gPLink")
 
strGPODisplayName = "Development Policy"
strGPOLinkOptions = 2
strNewGPLink = "[" & GetGPOADsPath & ";" & strGPOLinkOptions & "]"
 
objContainer.Put "gPLink", strExistingGPLink & strNewGPLink
objContainer.Put "gPOptions", "0"
 
objContainer.SetInfo
 
Function GetGPOADsPath
    Set objConnection = CreateObject("ADODB.Connection")  
    objConnection.Open "Provider=ADsDSOObject;"   
 
    Set objCommand = CreateObject("ADODB.Command")
    objCommand.ActiveConnection = objConnection
 
    objCommand.CommandText = _
      "<LDAP://cn=Policies,cn=System,dc=NA,dc=wisesoft,dc=co,dc=uk>;;" & _
          "distinguishedName,displayName;onelevel"
    Set objRecordSet = objCommand.Execute
 
    Do Until objRecordSet.EOF
        If objRecordSet.Fields("displayName") = strGPODisplayName Then
          GetGPOADsPath = "LDAP://" & objRecordSet.Fields("distinguishedName")
          objConnection.Close
          Exit Function
        End If
        objRecordSet.MoveNext
    Loop
    objConnection.Close
End Function
