'* Add SQL Server List Here...
SQLServers = array("SQL1","SQL2","SQL3")

FailedCount=0
SucceededCount=0

for i = 0 to ubound(SQLServers)
	html = html & getJobHistory(SQLServers(i))
next

displayHTML "<font size=6><b>Total Jobs Succeeded: " & SucceededCount & ", Total Jobs Failed: " & _
	     FailedCount & "</font></b><br><br>" & html, "SQL Server Agent History"


function getJobHistory(byval ServerName)
	set cmd = createobject("ADODB.Command")
	set cn = createobject("ADODB.Connection")
	set rs = createobject("ADODB.Recordset")

	cn.open "Provider=SQLOLEDB.1;Data Source=" & ServerName & ";Integrated Security=SSPI"

	cmd.activeconnection =cn
	cmd.commandtext = "select name, run_status, substring(cast(run_date as varchar(8)),7,2) + '/' + " & _
	"substring(cast(run_date as varchar(8)),5,2) + '/' + " & _
	"substring(cast(run_date as varchar(8)),1,4) + ' ' + case len(cast(run_time as varchar(6))) " & _
	"when 6 then left(cast(run_time as varchar(6)),2) + ':' + substring(cast(run_time as varchar(6)),3,2) " & _
	"when 5 then left(cast(run_time as varchar(6)),1) + ':' +  substring(cast(run_time as varchar(6)),2,2) " & _
	"else '00:00' end as 'RunTime' " & _
	"from msdb.dbo.sysjobhistory sjh " & _
	"join msdb.dbo.sysjobs_view sjv on sjh.job_id = sjv.job_id " & _
	"where instance_id in " & _
	"(select top 1 instance_id from msdb.dbo.sysjobhistory sjh2 where step_name = '(Job outcome)' and sjh.job_id " & _
	"= sjh2.job_id " & _
	"order by instance_id desc) " & _
	"order by run_status, instance_id DESC"
	
	set rs =cmd.execute
	message = "<B>SQL Server Agent Job History for " & ServerName & "</B><BR><BR>" & _
		  "<Table border =""1""><tr><td><b>Status</b></td><td><b>Execution Time</b></td><td><b>Job Name</b></td></B></tr>"

	while rs.eof<>true and rs.bof<>true
		message = message & "<font color='#0F00CD'><tr>"

		if rs(1) = 0 then
			result="FAILED"
			fontColour="#ff0000"
			FailedCount = FailedCount + 1
		elseif rs(1) = 1 then
			result="SUCCEEDED"
			fontColour="#0F00CD"
			SucceededCount = SucceededCount + 1
		else
			result="UNKNOWN"
			fontColour="#FF3300"
			FailedCount = FailedCount + 1
		end if

		message = message & "<td><font color='" & fontColour & "'>" & result & "</font></td><td><font color='" & fontColour & "'>" & _
					    rs(2) & "</font></td><td><font color='" & fontColour & "'>" & rs(0) & "</font></td>"
		rs.movenext
		message=message & "</tr></font>"
	wend
	message=message & "</Table><br>"
	'wscript.echo message ',64, "SQL Server Job History"
	getJobHistory = message
	cn.close

end function

sub displayHTML(byval html, Title)

	On Error Resume Next

	Set objExplorer = CreateObject("InternetExplorer.Application")

	objExplorer.Navigate "about:blank"
   	
	objExplorer.AddressBar = True
	objExplorer.MenuBar = True
	objExplorer.StatusBar = True
	objExplorer.ToolBar = True
	objExplorer.Visible = True

	objExplorer.Document.Title = Title
	objExplorer.Document.Body.InnerHTML = html

end sub