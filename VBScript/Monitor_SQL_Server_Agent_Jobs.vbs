ServerName = "(local)"

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
message = "SQL Agent Job History" & vbcrlf & vbcrlf & _
	  "Status" & vbtab & vbtab & "Execution Time" & vbtab & vbtab & "Job Name" & vbcrlf

while rs.eof<>true and rs.bof<>true
	if rs(1) = 0 then
		message = message & "FAILED" & vbtab & vbtab & rs(2) & vbtab & vbtab & rs(0) & vbcrlf
	elseif rs(1) = 1 then
		message = message & "SUCCEEDED" & vbtab & rs(2) & vbtab & vbtab & rs(0) & vbcrlf
	else
		message = message & "UNKNOWN" & vbtab & rs(2) & vbtab & vbtab & rs(0) & vbcrlf
	end if
	rs.movenext
wend

wscript.echo message ',64, "SQL Server Job History"