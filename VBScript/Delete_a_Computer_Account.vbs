strComputer = "computer-01"

set objComputer = GetObject("LDAP://CN=" & strComputer & _
    ",CN=Computers,DC=wisesoft,DC=co,DC=uk")
objComputer.DeleteObject (0)
