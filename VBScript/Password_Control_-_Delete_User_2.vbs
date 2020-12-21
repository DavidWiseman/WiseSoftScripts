ADsPath = wscript.arguments.named.item("ADsPath")

if ADsPath = "" then
	msgbox "Please type the name of the user to delete in the Password Control username textbox",vbOkOnly+vbExclamation,"Username required"
	wscript.quit
end if

set objUser = getobject(ADsPath)

result = msgbox("Are you sure you want to delete this user account?" & vbcrlf & _
	objUser.sAMAccountName,vbyesno+vbExclamation,"Confirm Delete")

if result = vbYes then
	set objContainer = getobject(objUser.Parent)
	objContainer.Delete "user","cn=" & objUser.cn
	msgbox "User account deleted successfully",vbOkOnly+vbInformation
end if