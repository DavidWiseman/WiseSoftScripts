Set objComputer = GetObject _
    ("LDAP://CN=computer-01,CN=Computers,DC=wisesoft,DC=co,DC=uk")

objComputer.SetPassword "computer-01$"
