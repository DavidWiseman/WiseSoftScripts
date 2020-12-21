const FileName ="domaincomputers.csv"
set cmd = createobject("ADODB.Command")
set cn = createobject("ADODB.Connection")
set rs = createobject("ADODB.Recordset")

cn.open "Provider=ADsDSOObject;"
cmd.activeconnection = cn

set objRoot = getobject("LDAP://RootDSE")

cmd.commandtext = "<LDAP://" & objRoot.get("defaultNamingContext") & ">;(objectCategory=Computer);" & _
		  "name,operatingsystem,operatingsystemservicepack, operatingsystemversion;subtree"
'**** Bypass 1000 record limitation ****
cmd.properties("page size")=1000

set rs = cmd.execute
set objFSO = createobject("Scripting.FileSystemObject")
set objCSV = objFSO.createtextfile(FileName)

q = """"

while rs.eof <> true and rs.bof <> true
	objcsv.writeline(q & rs("name") & q & "," &  q & rs("operatingsystem") & q & _
		"," & q & rs("operatingsystemservicepack") & _
		q & "," & q & rs("operatingsystemversion") & q)
	rs.movenext
wend

objCSV.Close
cn.close

wscript.echo "Finished"