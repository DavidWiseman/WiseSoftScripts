Set objComputer = GetObject _ 
    ("LDAP://CN=computer-01,CN=Computers,DC=wisesoft,DC=co,DC=uk")

objComputer.Put "Location" , "Building ABC, Room 123"
objComputer.SetInfo
