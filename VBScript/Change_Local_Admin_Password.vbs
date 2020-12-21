strPassword = "p@$$W0rD"
' Connect to local administrator user object
set objUser = getobject("WinNT://./Administrator,user")
' Change Password
objUser.SetPassword strPassword
objUser.SetInfo ' Save Changes