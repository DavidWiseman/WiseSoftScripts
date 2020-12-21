Const adInteger = 3
Const adParamInput = 1

' Prompt for user input (employeeId to use in SQL)
employeeId = inputbox("Please enter Employee ID")
if employeeId = "" then wscript.quit
if ISNUMERIC(employeeId) = false then
	wscript.echo "Invalid Input"
	wscript.quit
end if

' Connection String required to connect to MS Access database
connectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Program Files\Microsoft Office\OFFICE11\SAMPLES\Northwind.mdb;"
' SQL statement to run (with paremeter)
sql = "select EmployeeID,LastName from employees where EmployeeId = ?"

' Create ADO Connection/Command objects
set cn = createobject("ADODB.Connection")
set cmd = createobject("ADODB.Command")

' Open Connection
cn.open connectionString
' Associate connection object with command object
cmd.ActiveConnection = cn 
' Set the SQL statement of the command object
cmd.CommandText = sql

' Add parameter
cmd.parameters.append(cmd.createParameter("@p1", adInteger, adParamInput, , employeeId))

' Execute query
set rs = cmd.execute

' Check that query returned data
if rs.EOF<> True and rs.BOF<>True then
	wscript.echo "Last Name:" & rs("LastName")
else
	wscript.echo "Employee not found"
end if

' Close connection
cn.Close