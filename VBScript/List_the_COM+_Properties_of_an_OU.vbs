On Error Resume Next

Set objContainer = GetObject _
    ("LDAP://ou=Students,dc=wisesoft,dc=co,dc=uk")
 
strMsCOMUserPartitionSetLink = objContainer.Get("msCOM-UserPartitionSetLink")
WScript.Echo "ms-COMUserPartitionSetLink: " & strMsCOMUserPartitionSetLink
