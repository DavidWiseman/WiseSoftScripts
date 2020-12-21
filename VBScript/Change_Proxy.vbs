'* Example 1 : Change Proxy to IP Address 192.168.0.1 on port 8080
ChangeProxy "192.168.0.1:8080"

'* Example 2 : Change Proxy to a machine called isa1 on port 8080
ChangeProxy "isa1:8080"


sub ChangeProxy(Byval Proxy)

	set objShell = Wscript.CreateObject("Wscript.Shell")
	objShell.RegWrite "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Internet Settings\ProxyServer", Proxy, "REG_SZ"

end sub