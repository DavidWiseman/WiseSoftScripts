' Connection String required to connect to MS Access database
connectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Program Files\Microsoft Office\OFFICE11\SAMPLES\Northwind.mdb;"
' SQL statement to run
sql = "select EmployeeID,LastName from employees"

' Create ADO Connection/Command objects
set cn = createobject("ADODB.Connection")
set cmd = createobject("ADODB.Command")

' Open connection
cn.open connectionString
' Associate connection object with command object
cmd.ActiveConnection = cn
' Set the SQL statement of the command object
cmd.CommandText = sql

' Execute query
set rs = cmd.execute

' Enumerate each row in the result set
while rs.EOF <> true and rs.BOF <> True
	' Using ordinal
	wscript.echo "Employee ID: " & rs(0)
	' Using name
	wscript.echo "Last Name:" & rs("LastName")

	rs.movenext
wend

' Close Connection
cn.Close