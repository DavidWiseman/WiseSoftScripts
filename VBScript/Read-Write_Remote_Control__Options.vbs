dim objUser
'<<<< Bind to the user object using the distinguished name >>>>
set objUser = getobject("LDAP://cn=test.3,cn=users,dc=wisesoft,dc=co,dc=uk")

'<<<< Read >>>>
 
select case objuser.EnableRemoteControl
	case 0
		wscript.echo "Enable remote control: Disabled"
	case 1
		wscript.echo "Enable remote control: Enabled" & vbcrlf & _
		     	"Require user's permission: Enabled" & vbcrlf & _
			"Level of control: Interact with the session"
	case 2
		wscript.echo "Enable remote control: Enabled" & vbcrlf & _
		     	"Require user's permission: Disabled" & vbcrlf & _
			"Level of control: Interact with the session"
	case 3
		wscript.echo "Enable remote control: Enabled" & vbcrlf & _
		     	"Require user's permission: Enabled" & vbcrlf & _
			"Level of control: View the user's session"
	case 4
		wscript.echo "4Enable remote control: Enabled" & vbcrlf & _
		     	"Require user's permission: Disabled" & vbcrlf & _
			"Level of control: View the user's session"

end select


'<<<< Write >>>>

objuser.EnableRemoteControl = 1
objUser.setinfo	