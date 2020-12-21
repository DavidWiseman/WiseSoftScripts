const FileName = "userswithmailboxes.csv"

set cn = createobject("ADODB.Connection")
set cmd = createobject("ADODB.Command")
set rs = createobject("ADODB.Recordset")

cn.open "Provider=ADsDSOObject;"
cmd.activeconnection = cn

set objRoot = getobject("LDAP://RootDSE")

cmd.commandtext = "<LDAP://" & objRoot.get("defaultNamingContext") & _
		">;(&(objectClass=User)(HomeMDB=*));samaccountname,sn,givenname;subtree"
'**** Bypass the 1000 record limitation ****
cmd.properties("page size") = 1

set rs=cmd.execute

set objFSO = createobject("Scripting.FileSystemObject")
set objCSV = objFSO.createtextfile(FileName)
q = """"


while rs.eof <> true and rs.bof <> true
	objcsv.writeline(q & rs("samaccountname") &  q & "," & q  & _
			rs("sn") & q & "," & q & rs("givenName") & q)
	rs.movenext
wend

objCSV.Close
cn.close
	
wscript.echo "Finished"