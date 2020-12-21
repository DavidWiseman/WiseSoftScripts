option explicit
Const adOpenStatic = 3
Const adLockOptimistic = 3
Const adCmdText = &H0001
Const intWindowStyle = 7
Const blnWait = True
Dim strCSVFolder,strCSVFile,strNetDomParams
Dim objShell,cn,rs

' ************** Setup ************** 
' Folder where CSV File is located
' CSV file should have 1st field = oldname, 2nd field = newname with no header row
strCSVFolder = "C:\Temp\"
' CSV filename
strCSVFile = "test.csv"
' Additional parameters to pass to NetDom command
strNetDomParams = " /userd:DOMAIN\ADMINISTRATOR /passwordd:PASSWORD /usero:DOMAIN\ADMINISTRATOR /passwordo:PASSWORD /force "

'************************************ 

set objShell = wscript.createObject("wscript.shell")

' Setup ADO Connection to CSV file
Set cn = CreateObject("ADODB.Connection")
Set rs = CreateObject("ADODB.Recordset")

cn.Open "Provider=Microsoft.Jet.OLEDB.4.0;" & _
          "Data Source=" & strCSVFolder & ";" & _
          "Extended Properties=""text;HDR=NO;FMT=Delimited"""

rs.Open "SELECT * FROM [" & strCSVFile & "]", _
          cn, adOpenStatic, adLockOptimistic, adCmdText

do until rs.eof
	dim strOldName, strNewName, strCmd,intReturn
	strOldName = rs(0)
	strNewName = rs(1)
	strCmd = "cmd.exe /C netdom renamecomputer " & strOldName & " /newname:" & strNewName & strNetDomParams

	intReturn = objShell.Run(strCmd,intWindowStyle,blnWait)

	if intReturn = 0 then
		wscript.echo "Renamed '" & strOldName & "' to '" & strNewName & "'"
	else
		wscript.echo "Error renaming '" & strOldName & "' to '" & strNewName & "'" 
	end if

	rs.movenext
loop