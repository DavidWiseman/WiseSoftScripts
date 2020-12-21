Set SQLApp = CreateObject("SQLDMO.Application")

Set serverList = SQLApp.ListAvailableSQLServers

numServers = serverlist.count

For i = 1 To numServers
	wscript.echo serverList(i)
Next

Set SQLApp = Nothing