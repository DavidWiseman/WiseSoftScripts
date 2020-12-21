changeHomePage "http://www.wisesoft.co.uk"

Sub changeHomePage(byval URL)
	
	set objShell=Wscript.CreateObject("Wscript.Shell")
	objShell.RegWrite "HKEY_CURRENT_USER\Software\Microsoft\" & _
			   "Internet Explorer\Main\Start Page", URL, "REG_SZ"

End Sub	